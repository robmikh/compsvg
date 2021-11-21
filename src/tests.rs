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
            Direct2D::{
                Common::D2D_SIZE_F, ID2D1Device, ID2D1DeviceContext5, ID2D1Factory1,
                D2D1_DEVICE_CONTEXT_OPTIONS_NONE,
            },
            Direct3D11::{
                ID3D11Device, ID3D11DeviceContext, ID3D11Resource, ID3D11Texture2D, D3D11_MAP_READ,
                D3D11_TEXTURE2D_DESC,
            },
        },
        System::WinRT::{RoInitialize, RO_INIT_MULTITHREADED},
    },
    UI::Composition::{Compositor, Core::CompositorController, Visual},
};

use crate::{convert_svg_document_to_composition_shapes, SvgCompositionShapes};

struct TestRuntime {
    pub _controller: DispatcherQueueController,
    pub compositor_controller: CompositorController,
    pub compositor: Compositor,

    pub d3d_device: ID3D11Device,
    pub d3d_context: ID3D11DeviceContext,
    pub _d2d_factory: ID2D1Factory1,
    pub _d2d_device: ID2D1Device,
    pub d2d_context: ID2D1DeviceContext5,
}

impl TestRuntime {
    pub fn new() -> Result<Self> {
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

        Ok(Self {
            _controller,
            compositor_controller,
            compositor,

            d3d_device,
            d3d_context,
            _d2d_factory: d2d_factory,
            _d2d_device: d2d_device,
            d2d_context,
        })
    }

    pub fn load_svg(&self, folder: &str, file_stem: &str) -> Result<SvgCompositionShapes> {
        let card_faces_path = {
            let mut path = std::env::current_dir().unwrap();
            path.push(format!("test_data\\{}", folder));
            path
        };
        let folder =
            StorageFolder::GetFolderFromPathAsync(card_faces_path.to_str().unwrap())?.get()?;
        let file = folder.GetFileAsync(format!("{}.svg", file_stem))?.get()?;
        let stream = file.OpenReadAsync()?.get()?;

        let viewport = D2D_SIZE_F {
            width: 1.0,
            height: 1.0,
        };
        let document = unsafe {
            self.d2d_context
                .CreateSvgDocument(stream.as_istream()?, &viewport)?
        };
        let shape_info = convert_svg_document_to_composition_shapes(&self.compositor, &document)?;
        Ok(shape_info)
    }

    pub fn capture_visual(&self, visual: &Visual) -> Result<ID3D11Texture2D> {
        let item = GraphicsCaptureItem::CreateFromVisual(visual)?;
        let direct3d_device = create_direct3d_device(&self.d3d_device)?;
        let frame_pool = Direct3D11CaptureFramePool::CreateFreeThreaded(
            direct3d_device,
            DirectXPixelFormat::B8G8R8A8UIntNormalized,
            1,
            item.Size()?,
        )?;
        let session = frame_pool.CreateCaptureSession(item)?;

        let (sender, receiver) = channel();
        frame_pool.FrameArrived(
            TypedEventHandler::<Direct3D11CaptureFramePool, IInspectable>::new(
                move |frame_pool, _| {
                    let frame_pool = frame_pool.as_ref().unwrap();
                    let frame = frame_pool.TryGetNextFrame()?;
                    sender.send(frame).unwrap();
                    Ok(())
                },
            ),
        )?;
        session.StartCapture()?;
        self.compositor_controller.Commit()?;
        let capture_frame = receiver.recv().unwrap();
        session.Close()?;
        frame_pool.Close()?;
        let texture: ID3D11Texture2D = get_d3d_interface_from_object(&capture_frame.Surface()?)?;
        Ok(texture)
    }

    pub fn get_texture_bytes(&self, texture: &ID3D11Texture2D) -> Result<(Vec<u8>, u32, u32)> {
        unsafe {
            let texture = copy_texture(&self.d3d_device, &self.d3d_context, &texture, true)?;
            let mut desc = D3D11_TEXTURE2D_DESC::default();
            texture.GetDesc(&mut desc as *mut _);

            let resource: ID3D11Resource = texture.cast()?;
            let mapped = self
                .d3d_context
                .Map(Some(resource.clone()), 0, D3D11_MAP_READ, 0)?;

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

            self.d3d_context.Unmap(Some(resource), 0);

            Ok((bytes, desc.Width, desc.Height))
        }
    }

    pub fn save_texture_as_png(
        &self,
        folder: &str,
        file_stem: &str,
        texture: &ID3D11Texture2D,
    ) -> Result<()> {
        let (bytes, width, height) = self.get_texture_bytes(texture)?;
        let output_path = {
            let mut path = std::env::current_dir().unwrap();
            path.push(format!("test_output\\{}", folder));
            path
        };
        if !output_path.exists() {
            std::fs::create_dir_all(&output_path).unwrap();
        }
        let folder = StorageFolder::GetFolderFromPathAsync(output_path.to_str().unwrap())?.get()?;
        let file = folder
            .CreateFileAsync(
                format!("{}.png", file_stem),
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
        Ok(())
    }
}

fn dump_card_to_png(runtime: &TestRuntime, card_name: &str) -> Result<()> {
    let folder = "card_faces";

    // Load the king of diamonds
    let shape_info = runtime.load_svg(folder, card_name)?;

    // Setup a visual tree for the card
    const CARD_SIZE: Vector2 = Vector2 { X: 167.0, Y: 243.0 };
    const SCALE: f32 = 5.0;
    let root = runtime.compositor.CreateContainerVisual()?;
    root.SetSize(CARD_SIZE * SCALE)?;
    let shape_visual = runtime.compositor.CreateShapeVisual()?;
    let shape_container = runtime.compositor.CreateContainerShape()?;
    shape_visual.Shapes()?.Append(&shape_container)?;
    shape_visual.SetSize(CARD_SIZE)?;
    shape_visual.SetScale(Vector3::new(SCALE, SCALE, 1.0))?;
    shape_container.Shapes()?.Append(&shape_info.root_shape)?;
    shape_visual.SetViewBox(&shape_info.view_box)?;
    root.Children()?.InsertAtTop(shape_visual)?;

    runtime.compositor_controller.Commit()?;

    // Capture the visual tree
    let source_texture = runtime.capture_visual(&root.cast()?)?;

    // Encode the image to disk
    runtime.save_texture_as_png(folder, card_name, &source_texture)?;

    Ok(())
}

#[test]
fn king_of_diamonds() -> Result<()> {
    unsafe { RoInitialize(RO_INIT_MULTITHREADED)? };
    let runtime = TestRuntime::new()?;
    dump_card_to_png(&runtime, "king_of_diamonds")?;
    Ok(())
}
