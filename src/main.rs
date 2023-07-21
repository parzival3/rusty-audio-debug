use windows::{core::*, Win32::System::Com::*, Win32::Media::Audio::*, Win32::UI::Shell::PropertiesSystem::*};

unsafe fn pwstr_to_string(string: PWSTR) -> String {
    let mut end = string.0;
    while *end != 0 {
        end = end.add(1);
    }
    let string_id = String::from_utf16_lossy(std::slice::from_raw_parts(
        string.0,
        end.offset_from(string.0) as _,
    ));
    return string_id;
}

unsafe fn u16_to_string(string: [u16; 200]) -> String {
    let mut end = 0;
    while string[end] != 0 {
        end = end + 1;
    }
    let string_id = String::from_utf16_lossy(std::slice::from_raw_parts(
        &string[0],
        end
    ));
    return string_id;
}

fn main() -> Result<()> {
    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED)?;
        let imm_device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_INPROC_SERVER)
                ?;
        let imm_device_collection: IMMDeviceCollection = imm_device_enumerator.EnumAudioEndpoints(
            eAll,
          DEVICE_STATE_ACTIVE,
        )?;

        let device_count = imm_device_collection.GetCount()?;
        println!("Device count: {}", device_count);
        for i in 0..device_count {
            let imm_device = imm_device_collection.Item(i)?;
            let device_id = imm_device.GetId()?;
            println!("Device ID: {:?}", device_id);
            let imm_property_store = imm_device.OpenPropertyStore(STGM_READ)?;
            let count : u32 = imm_property_store.GetCount().expect("Couldn't get the number of properties");
            for p_indx in 0..count {
                let mut prop: PROPERTYKEY = Default::default(); 
                imm_property_store.GetAt(p_indx, &mut prop)?;
                let mut the_string : [u16; 200] = [0; 200];
                PSStringFromPropertyKey(&prop, &mut the_string).expect("Couldn't get the string for this property key");
                let value = imm_property_store.GetValue(&prop).expect("Couldn't get the value at index");
                let pwstr_value = PropVariantToStringAlloc(&value).expect("Couldn't convert to PWSTR");
                let string_value = pwstr_to_string(pwstr_value);
                let name = PSGetNameFromPropertyKey(&prop);
                let string_name  = match name {
                    Ok(pwstr_name) => pwstr_to_string(pwstr_name),
                    Err(_) => String::new()
                };
        
                println!("\t Property '{}': '{}' GUIID '{}'", string_name, string_value, u16_to_string(the_string));
                let mut pinterface: *mut std::ffi::c_void = std::ptr::null_mut();
                let res = PSGetPropertyDescription(&prop,&IPropertyDescription::IID, &mut pinterface as *mut _);
                match res {
                    Ok(_) =>  {
                        let property_description = std::mem::transmute::<*mut std::ffi::c_void, IPropertyDescription>(pinterface);
                        let res = property_description.GetDisplayName();
                        if let Ok(pwstr_name) = res { 
                                 println!("\t\t Display Name: {:?}", pwstr_to_string(pwstr_name));
                                 println!("------------------------------------------------------------------------------------------------------------------------------------------------------------------------"); 
                        }
                    },
                    Err(_) => println!("This property doesn't have a description")
                }
        
            }
        }
    }
    println!("Hello, world!");
    Ok(())
}
