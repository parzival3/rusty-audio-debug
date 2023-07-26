use windows::{
    core::*, Win32::Devices::FunctionDiscovery::PKEY_Device_InstanceId, Win32::Media::Audio::*,
    Win32::System::Com::*, Win32::UI::Shell::PropertiesSystem::*,
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
    audio_toplogy_id: String,
    properties: Vec<DeviceProperty>,
}

pub const PKEY_ISS_USB_DEVICE: PROPERTYKEY = PROPERTYKEY {
    fmtid: ::windows::core::GUID::from_u128(0xB3F8FA53_0004_438E_9003_51A46E139BFC),
    pid: 39,
};

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

fn enumerate_imm_devices(imm_device_enumerator: IMMDeviceEnumerator) -> Result<Vec<AudioDevice>> {
    let device_count: u32;
    let imm_device_collection: IMMDeviceCollection;

    unsafe {
        imm_device_collection =
            imm_device_enumerator.EnumAudioEndpoints(eAll, DEVICE_STATE_ACTIVE)?;
        device_count = imm_device_collection.GetCount()?;
    }

    let mut audio_devices = Vec::new();
    for i in 0..device_count {
        let imm_device = unsafe { imm_device_collection.Item(i)? };
        let imm_property_store = unsafe { imm_device.OpenPropertyStore(STGM_READ)? };
        let device_id = unsafe { imm_device.GetId()? };
        let device_properties = collect_device_properties(&imm_device)?;
        let audio_controller = match get_audio_device(
            &unsafe { get_audio_controller_id(imm_device)? },
            &imm_device_enumerator,
        ) {
            Ok(id) => id,
            Err(e) => format!("Error {e}"),
        };

        let toploogy = unsafe { imm_property_store.GetValue(&PKEY_ISS_USB_DEVICE)? };

        audio_devices.push(AudioDevice {
            id: unsafe { device_id.to_string()? },
            audio_toplogy_id: unsafe { PropVariantToStringAlloc(&toploogy)?.to_string()? },
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

fn get_audio_device(device_id: &String, enumerator: &IMMDeviceEnumerator) -> Result<String> {
    // TODO: This is a hack, we should be able to use the device_id directly
    let mut text = device_id.as_str()[..device_id.find("/").unwrap_or(device_id.len())]
        .encode_utf16()
        .collect::<Vec<_>>();
    text.push(0);
    let wstr = PCWSTR::from_raw(text.as_ptr());
    let device: IMMDevice = unsafe { enumerator.GetDevice(wstr)? };
    let property_store: IPropertyStore = unsafe { device.OpenPropertyStore(STGM_READ)? };
    let value = unsafe { property_store.GetValue(&PKEY_Device_InstanceId)? };
    unsafe {
        PropVariantToStringAlloc(&value)?
            .to_string()
            .map_err(|e| e.into())
    }
}

fn collect_device_properties(device: &IMMDevice) -> Result<Vec<DeviceProperty>> {
    let imm_property_store = unsafe { device.OpenPropertyStore(STGM_READ)? };
    let number_of_properties: u32 = unsafe { imm_property_store.GetCount()? };
    let mut device_properties = Vec::new();
    for property_index in 0..number_of_properties {
        device_properties.push(collect_property(&imm_property_store, property_index)?);
    }
    Ok(device_properties)
}

fn collect_property(
    imm_property_store: &IPropertyStore,
    property_index: u32,
) -> Result<DeviceProperty> {
    let mut prop: PROPERTYKEY = Default::default();
    unsafe { imm_property_store.GetAt(property_index, &mut prop)? };
    let mut raw_guid_string: [u16; 200] = [0; 200];
    unsafe { PSStringFromPropertyKey(&prop, &mut raw_guid_string)? };
    let value = unsafe { imm_property_store.GetValue(&prop)? };
    let value = unsafe { PropVariantToStringAlloc(&value)?.to_string()? };

    let mut description = String::new();
    let name = if let Ok(name_pwstr) = unsafe { PSGetNameFromPropertyKey(&prop) } {
        unsafe { name_pwstr.to_string()? }
    } else {
        String::new()
    };

    let mut pinterface: *mut std::ffi::c_void = std::ptr::null_mut();
    let res = unsafe {
        PSGetPropertyDescription(&prop, &IPropertyDescription::IID, &mut pinterface as *mut _)
    };

    if res.is_ok() {
        let property_description = unsafe {
            std::mem::transmute::<*mut std::ffi::c_void, IPropertyDescription>(pinterface)
        };
        let res = unsafe { property_description.GetDisplayName() };
        if let Ok(property_description) = res {
            description = unsafe { property_description.to_string()? };
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
    let audio_devices: Vec<AudioDevice>;
    unsafe {
        let imm_device_enumerator = get_imm_device_enumerator()?;
        audio_devices = enumerate_imm_devices(imm_device_enumerator)?;
    }

    for audio_device in audio_devices {
        println!("Device ID: {}", audio_device.id);
        println!("Audio Controller ID: {}", audio_device.audio_controller_id);
        println!("Audio Intel Smart Sound USB device: {}", audio_device.audio_toplogy_id);
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
    Ok(())
}
