use windows::{
    core::*, Win32::Media::Audio::*, Win32::System::Com::*, Win32::UI::Shell::PropertiesSystem::*,
    Win32::Devices::FunctionDiscovery::PKEY_Device_InstanceId
};

struct DeviceProperty {
    name: String,
    value: String,
    guid: String,
    description: String,
}

struct AudioDevice {
    id: String,
    audio_controller_id: String,
    properties: Vec<DeviceProperty>,
}

fn u16_to_string(string: [u16; 200]) -> String {
    let mut end = 0;
    while string[end] != 0 {
        end = end + 1;
    }
    unsafe { String::from_utf16_lossy(std::slice::from_raw_parts(&string[0], end)) }
}

unsafe fn get_imm_device_enumerator() -> Result<IMMDeviceEnumerator> {
    CoInitializeEx(None, COINIT_APARTMENTTHREADED)?;
    let imm_device_enumerator: IMMDeviceEnumerator =
        CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_INPROC_SERVER)?;
    Ok(imm_device_enumerator)
}

unsafe fn enumerate_imm_devices(
    imm_device_enumerator: IMMDeviceEnumerator,
) -> Result<Vec<AudioDevice>> {
    let imm_device_collection: IMMDeviceCollection =
        imm_device_enumerator.EnumAudioEndpoints(eAll, DEVICE_STATE_ACTIVE)?;
    let device_count = imm_device_collection.GetCount()?;
    let mut audio_devices = Vec::new();
    for i in 0..device_count {
        let imm_device = imm_device_collection.Item(i)?;
        let imm_property_store = imm_device.OpenPropertyStore(STGM_READ)?;
        let number_of_properties: u32 = imm_property_store.GetCount()?;
        let device_id = imm_device.GetId()?;
        let mut device_properties = Vec::new();
        for property_index in 0..number_of_properties {
            device_properties.push(collect_property(&imm_property_store, property_index)?);
        }
        let audio_controller = match get_audio_device(&get_audio_controller_id(imm_device)?, &imm_device_enumerator) {
            Ok(id) => id,
            Err(e) => format!("Error {e}"),
        };

        audio_devices.push(AudioDevice {
            id: device_id.to_string()?,
            audio_controller_id: audio_controller,
            properties: device_properties,
        });
    }
    Ok(audio_devices)
}

unsafe fn get_audio_controller_id(device: IMMDevice) -> Result<String> {
    let topology: IDeviceTopology = device.Activate(CLSCTX_ALL, None)?;
    let connector: IConnector = topology.GetConnector(0)?;
    connector
        .GetConnectorIdConnectedTo()?
        .to_string()
        .map_err(|e| e.into())
}

unsafe fn get_audio_device(device_id: &String, enumerator: &IMMDeviceEnumerator) -> Result<String> {
    let mut text = device_id.as_str()[..device_id.find("\\global").expect("No global") + "\\global".len()].encode_utf16().collect::<Vec<_>>();
    text.push(0);
    let wstr = PCWSTR::from_raw(text.as_ptr());
    println!("wstr: {}", wstr.display());
    // let wstr = w!("{2}.\\\\?\\usb#vid_1395&pid_029e&mi_00#a&17daf453&0&0000#{6994ad04-93ef-11d0-a3cc-00a0c9223196}\\global");
    println!("wstr: {}", wstr.display());
    let device: IMMDevice = enumerator.GetDevice(wstr)?;
    let property_store: IPropertyStore = device.OpenPropertyStore(STGM_READ)?;
    let value = property_store.GetValue(&PKEY_Device_InstanceId)?;
    PropVariantToStringAlloc(&value)?.to_string().map_err(|e| e.into())
}

unsafe fn collect_property(
    imm_property_store: &IPropertyStore,
    property_index: u32,
) -> Result<DeviceProperty> {
    let mut prop: PROPERTYKEY = Default::default();
    imm_property_store.GetAt(property_index, &mut prop)?;
    let mut raw_guid_string: [u16; 200] = [0; 200];
    PSStringFromPropertyKey(&prop, &mut raw_guid_string)?;
    let value = imm_property_store.GetValue(&prop)?;
    let value = PropVariantToStringAlloc(&value)?.to_string()?;

    let mut description = String::new();
    let name = if let Ok(name_pwstr) = PSGetNameFromPropertyKey(&prop) {
        name_pwstr.to_string()?
    } else {
        String::new()
    };

    let mut pinterface: *mut std::ffi::c_void = std::ptr::null_mut();
    let res =
        PSGetPropertyDescription(&prop, &IPropertyDescription::IID, &mut pinterface as *mut _);

    if res.is_ok() {
        let property_description =
            std::mem::transmute::<*mut std::ffi::c_void, IPropertyDescription>(pinterface);
        let res = property_description.GetDisplayName();
        if let Ok(property_description) = res {
            description = property_description.to_string()?;
        }
    }

    Ok(DeviceProperty {
        name,
        value,
        guid: u16_to_string(raw_guid_string),
        description,
    })
}

fn main() -> Result<()> {
    unsafe {
        let imm_device_enumerator = get_imm_device_enumerator()?;
        let audio_devices = enumerate_imm_devices(imm_device_enumerator)?;
        for audio_device in audio_devices {
            println!("Device ID: {}", audio_device.id);
            println!("Audio Controller ID: {}", audio_device.audio_controller_id);
            for property in audio_device.properties {
                println!(
                    "\t Property '{}'
                     \t\t Value '{}'
                     \t\t GUIID '{}'
                     \t\t Description '{}'",
                    property.name, property.value, property.guid, property.description
                );
            }
        }
    }
    Ok(())
}
