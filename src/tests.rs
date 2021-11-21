use std::sync::mpsc::channel;

use robmikh_common::{
    desktop::dispatcher_queue::DispatcherQueueControllerExtensions,
    universal::{
        d2d::{create_d2d_device, create_d2d_factory},
        d3d::{
            copy_texture, create_d3d_device, create_direct3d_device, get_d3d_interface_from_object,
        },
        stream::AsIStream,
    },
};
use windows::{
    core::{IInspectable, Interface, Result},
    Foundation::{
        Numerics::{Vector2, Vector3},
        TypedEventHandler,
    },
    Graphics::{
        Capture::{Direct3D11CaptureFramePool, GraphicsCaptureItem},
        DirectX::DirectXPixelFormat,
        Imaging::{BitmapAlphaMode, BitmapEncoder, BitmapPixelFormat},
    },
    Storage::{CreationCollisionOption, FileAccessMode, StorageFolder},
    System::DispatcherQueueController,
    Win32::{
        Graphics::{
            Direct2D::{Common::D2D_SIZE_F, ID2D1DeviceContext5, D2D1_DEVICE_CONTEXT_OPTIONS_NONE},
            Direct3D11::{ID3D11Resource, ID3D11Texture2D, D3D11_MAP_READ, D3D11_TEXTURE2D_DESC},
        },
        System::WinRT::{RoInitialize, RO_INIT_MULTITHREADED},
    },
    UI::Composition::{CompositionBackfaceVisibility, Core::CompositorController},
};

use crate::convert_svg_document_to_composition_shapes;

#[test]
fn king_of_diamonds() -> Result<()> {
    unsafe { RoInitialize(RO_INIT_MULTITHREADED)? };

    // We'll need this for paths later
    let current_dir = std::env::current_dir().unwrap();

    // Setup our compositor so that we can capture the svg output
    // with Windows.Graphics.Capture later.
    let _controller =
        DispatcherQueueController::create_dispatcher_queue_controller_for_current_thread()?;
    let compositor_controller = CompositorController::new()?;
    let compositor = compositor_controller.Compositor()?;

    // Setup d2d so we can load the svg document
    let d3d_device = create_d3d_device()?;
    let d3d_context = {
        let mut d3d_context = None;
        unsafe { d3d_device.GetImmediateContext(&mut d3d_context) };
        d3d_context.unwrap()
    };
    let d2d_factory = create_d2d_factory()?;
    let d2d_device = create_d2d_device(&d2d_factory, &d3d_device)?;
    let d2d_context: ID2D1DeviceContext5 = unsafe {
        d2d_device
            .CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE)?
            .cast()?
    };

    // Load the king of diamonds
    let shape_info = {
        let card_faces_path = {
            let mut path = current_dir.clone();
            path.push("test_data\\card_faces");
            path
        };
        let folder =
            StorageFolder::GetFolderFromPathAsync(card_faces_path.to_str().unwrap())?.get()?;
        let file = folder.GetFileAsync("king_of_diamonds.svg")?.get()?;
        let stream = file.OpenReadAsync()?.get()?;

        let viewport = D2D_SIZE_F {
            width: 1.0,
            height: 1.0,
        };
        let document = unsafe { d2d_context.CreateSvgDocument(stream.as_istream()?, &viewport)? };
        let shape_info = convert_svg_document_to_composition_shapes(&compositor, &document)?;
        shape_info
    };

    // Setup a visual tree for the card
    const CARD_SIZE: Vector2 = Vector2 { X: 167.0, Y: 243.0 };
    const SCALE: f32 = 5.0;
    let root = compositor.CreateContainerVisual()?;
    root.SetSize(CARD_SIZE * SCALE)?;
    let shape_visual = compositor.CreateShapeVisual()?;
    let shape_container = compositor.CreateContainerShape()?;
    shape_visual.Shapes()?.Append(&shape_container)?;
    shape_visual.SetSize(CARD_SIZE)?;
    shape_visual.SetScale(Vector3::new(SCALE, SCALE, 1.0))?;
    shape_visual.SetBackfaceVisibility(CompositionBackfaceVisibility::Hidden)?;
    shape_container.Shapes()?.Append(&shape_info.root_shape)?;
    shape_visual.SetViewBox(&shape_info.view_box)?;
    root.Children()?.InsertAtTop(shape_visual)?;

    compositor_controller.Commit()?;

    // Capture the visual tree
    let item = GraphicsCaptureItem::CreateFromVisual(root)?;
    let direct3d_device = create_direct3d_device(&d3d_device)?;
    let frame_pool = Direct3D11CaptureFramePool::CreateFreeThreaded(
        direct3d_device,
        DirectXPixelFormat::B8G8R8A8UIntNormalized,
        1,
        item.Size()?,
    )?;
    let session = frame_pool.CreateCaptureSession(item)?;

    let (sender, receiver) = channel();
    frame_pool.FrameArrived(
        TypedEventHandler::<Direct3D11CaptureFramePool, IInspectable>::new(move |frame_pool, _| {
            let frame_pool = frame_pool.as_ref().unwrap();
            let frame = frame_pool.TryGetNextFrame()?;
            sender.send(frame).unwrap();
            Ok(())
        }),
    )?;
    session.StartCapture()?;
    compositor_controller.Commit()?;
    let capture_frame = receiver.recv().unwrap();
    session.Close()?;
    frame_pool.Close()?;

    // Retrieve the raw bytes
    let (bytes, width, height) = unsafe {
        let source_texture: ID3D11Texture2D =
            get_d3d_interface_from_object(&capture_frame.Surface()?)?;
        let texture = copy_texture(&d3d_device, &d3d_context, &source_texture, true)?;
        let mut desc = D3D11_TEXTURE2D_DESC::default();
        texture.GetDesc(&mut desc as *mut _);

        let resource: ID3D11Resource = texture.cast()?;
        let mapped = d3d_context.Map(Some(resource.clone()), 0, D3D11_MAP_READ, 0)?;

        // Get a slice of bytes
        let slice: &[u8] = {
            std::slice::from_raw_parts(
                mapped.pData as *const _,
                (desc.Height * mapped.RowPitch) as usize,
            )
        };

        let bytes_per_pixel = 4;
        let mut bytes = vec![0u8; (desc.Width * desc.Height * bytes_per_pixel) as usize];
        for row in 0..desc.Height {
            let data_begin = (row * (desc.Width * bytes_per_pixel)) as usize;
            let data_end = ((row + 1) * (desc.Width * bytes_per_pixel)) as usize;
            let slice_begin = (row * mapped.RowPitch) as usize;
            let slice_end = slice_begin + (desc.Width * bytes_per_pixel) as usize;
            bytes[data_begin..data_end].copy_from_slice(&slice[slice_begin..slice_end]);
        }

        d3d_context.Unmap(Some(resource), 0);

        (bytes, desc.Width, desc.Height)
    };

    // Encode the image to disk
    {
        let output_path = {
            let mut path = current_dir.clone();
            path.push("test_output\\card_faces");
            path
        };
        if !output_path.exists() {
            std::fs::create_dir_all(&output_path).unwrap();
        }
        let folder = StorageFolder::GetFolderFromPathAsync(output_path.to_str().unwrap())?.get()?;
        let file = folder
            .CreateFileAsync(
                "king_of_diamonds.png",
                CreationCollisionOption::ReplaceExisting,
            )?
            .get()?;
        let stream = file.OpenAsync(FileAccessMode::ReadWrite)?.get()?;
        let encoder = BitmapEncoder::CreateAsync(BitmapEncoder::PngEncoderId()?, stream)?.get()?;
        encoder.SetPixelData(
            BitmapPixelFormat::Bgra8,
            BitmapAlphaMode::Premultiplied,
            width,
            height,
            1.0,
            1.0,
            &bytes,
        )?;
        encoder.FlushAsync()?.get()?;
    }

    Ok(())
}
