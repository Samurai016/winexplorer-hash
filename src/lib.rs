#![allow(non_snake_case)]

use std::ffi::c_void;
use std::sync::atomic::{AtomicU32, Ordering};

use windows::{
    core::{implement, Result, w, GUID, HRESULT, IUnknown, PCWSTR, PWSTR, ComInterface},
    Win32::{
        Foundation::{
            BOOL, CLASS_E_CLASSNOTAVAILABLE, E_FAIL, E_INVALIDARG, E_UNEXPECTED, MAX_PATH, S_FALSE, S_OK,
            HMODULE,
        },
        System::{
            Com::{
                IClassFactory, IClassFactory_Impl, IStream,
                STATFLAG_NONAME, CoTaskMemAlloc,
                StructuredStorage::PROPVARIANT,
            },
            LibraryLoader::GetModuleFileNameW,
            Registry::{
                RegCloseKey, RegCreateKeyExW, RegDeleteTreeW, RegSetValueExW, HKEY, HKEY_CLASSES_ROOT,
                KEY_ALL_ACCESS, REG_OPTION_NON_VOLATILE, REG_SZ,
                RegOpenKeyExW, RegQueryValueExW, HKEY_CURRENT_USER, KEY_READ, REG_QWORD, REG_DWORD,
            },
        },
        UI::Shell::PropertiesSystem::{
            IInitializeWithStream, IInitializeWithStream_Impl,
            IPropertyStore, IPropertyStore_Impl, PSRegisterPropertySchema, PSUnregisterPropertySchema,
            PROPERTYKEY,
        },
        System::Variant::VT_LPWSTR,
    },
};

const CLSID_HASH_PROPERTY_HANDLER: GUID = GUID::from_u128(0x8E97E8B8_5A24_4FDB_AA9D_9F319BE24B02);

const  PKEY_FILE_HASH_MD5: PROPERTYKEY = PROPERTYKEY {
    fmtid: GUID::from_u128(0xF8BFA532_6D39_44DB_9EFE_DEFECC584EBC),
    pid: 100,
};

static DLL_REF_COUNT: AtomicU32 = AtomicU32::new(0);
static mut DLL_INSTANCE: HMODULE = HMODULE(0);

#[no_mangle]
pub extern "system" fn DllMain(
    inst: HMODULE,
    reason: u32,
    _reserved: *const c_void,
) -> BOOL {
    if reason == 1 /* DLL_PROCESS_ATTACH */ {
        unsafe {
            DLL_INSTANCE = inst;
        }
    }
    BOOL::from(true)
}

#[implement(IInitializeWithStream, IPropertyStore)]
struct HashPropertyHandler {
    hash_value: std::sync::RwLock<Option<String>>,
}

impl HashPropertyHandler {
    fn new() -> Self {
        DLL_REF_COUNT.fetch_add(1, Ordering::SeqCst);
        Self {
            hash_value: std::sync::RwLock::new(None),
        }
    }
}

impl Drop for HashPropertyHandler {
    fn drop(&mut self) {
        DLL_REF_COUNT.fetch_sub(1, Ordering::SeqCst);
    }
}

fn get_max_file_size() -> u64 {
    let mut hkey = HKEY::default();
    let status = unsafe {
        RegOpenKeyExW(
            HKEY_CURRENT_USER,
            w!("Software\\WinExplorerHash"),
            0,
            KEY_READ,
            &mut hkey,
        )
    };

    if status.is_ok() {
        let mut data = [0u8; 8];
        let mut data_size = 8u32;
        let mut reg_type = 0u32;
        
        let query_status = unsafe {
            RegQueryValueExW(
                hkey,
                w!("MaxFileSizeBytes"),
                None,
                Some(&mut reg_type),
                Some(data.as_mut_ptr()),
                Some(&mut data_size),
            )
        };

        unsafe { let _ = RegCloseKey(hkey); }

        if query_status.is_ok() {
            if reg_type == REG_QWORD.0 && data_size == 8 {
                let bytes: [u8; 8] = data.try_into().unwrap();
                return u64::from_ne_bytes(bytes);
            } else if reg_type == REG_DWORD.0 && data_size == 4 {
                let bytes: [u8; 4] = data[0..4].try_into().unwrap();
                return u32::from_ne_bytes(bytes) as u64;
            }
        }
    }

    10 * 1024 * 1024 // 10 MB default
}

impl IInitializeWithStream_Impl for HashPropertyHandler {
    fn Initialize(&self, pstream: Option<&IStream>, _grfmode: u32) -> Result<()> {
        let stream = match pstream {
            Some(s) => s,
            None => {
                return Err(windows::core::Error::from(E_INVALIDARG));
            }
        };

        let mut stat_stg = Default::default();
        let size = if unsafe { stream.Stat(&mut stat_stg, STATFLAG_NONAME) }.is_ok() {
            stat_stg.cbSize
        } else {
            0
        };

        if size > get_max_file_size() {
            *self.hash_value.write().unwrap() = Some("> 10MB (Skipped)".to_string());
            return Ok(());
        }

        let mut hasher = md5::Md5::new();
        use md5::Digest;
        
        // 32 KB buffer
        let mut buffer = [0u8; 32768];

        loop {
            let mut bytes_read = 0u32;
            unsafe {
                let _ = stream.Read(
                    buffer.as_mut_ptr() as *mut c_void,
                    buffer.len() as u32,
                    Some(&mut bytes_read),
                );
            }
            if bytes_read == 0 {
                break;
            }
            hasher.update(&buffer[..bytes_read as usize]);
        }

        let result = hasher.finalize();
        let hash_hex = hex::encode(result);

        *self.hash_value.write().unwrap() = Some(hash_hex);
        Ok(())
    }
}

impl IPropertyStore_Impl for HashPropertyHandler {
    fn GetCount(&self) -> Result<u32> {
        Ok(1)
    }

    fn GetAt(&self, iprop: u32, pkey: *mut PROPERTYKEY) -> Result<()> {
        if iprop == 0 {
            unsafe { *pkey =  PKEY_FILE_HASH_MD5 };
            Ok(())
        } else {
            Err(windows::core::Error::from(E_INVALIDARG))
        }
    }

    fn GetValue(&self, key: *const PROPERTYKEY) -> Result<PROPVARIANT> {
        let key = unsafe { *key };
        if key.fmtid ==  PKEY_FILE_HASH_MD5.fmtid && key.pid ==  PKEY_FILE_HASH_MD5.pid {
            let val = self.hash_value.read().unwrap();
            match &*val {
                Some(s) => unsafe {
                    let mut pv = PROPVARIANT::default();
                    let wide: Vec<u16> = s.encode_utf16().chain(std::iter::once(0)).collect();
                    let bytes = wide.len() * 2;
                    let ptr = CoTaskMemAlloc(bytes);
                    if !ptr.is_null() {
                        std::ptr::copy_nonoverlapping(wide.as_ptr() as *const u8, ptr as *mut u8, bytes);
                        (*pv.Anonymous.Anonymous).vt = windows::Win32::System::Variant::VARENUM(VT_LPWSTR.0 as u16);
                        (*pv.Anonymous.Anonymous).Anonymous.pwszVal = PWSTR(ptr as *mut u16);
                    }
                    Ok(pv)
                },
                None => {
                    Ok(PROPVARIANT::default())
                },
            }
        } else {
            Ok(PROPVARIANT::default())
        }
    }

    fn SetValue(&self, _key: *const PROPERTYKEY, _propvar: *const PROPVARIANT) -> Result<()> {
        Err(windows::core::Error::from(E_UNEXPECTED))
    }

    fn Commit(&self) -> Result<()> {
        Ok(())
    }
}

#[implement(IClassFactory)]
struct Factory;

impl IClassFactory_Impl for Factory {
    fn CreateInstance(
        &self,
        outer: Option<&IUnknown>,
        riid: *const GUID,
        ppv: *mut *mut c_void,
    ) -> Result<()> {
        unsafe { *ppv = std::ptr::null_mut() };
        if outer.is_some() {
            return Err(windows::core::Error::from(CLASS_E_CLASSNOTAVAILABLE));
        }

        let instance = HashPropertyHandler::new();
        let unknown: IUnknown = instance.into();
        unsafe { unknown.query(&*riid, ppv).ok() }
    }

    fn LockServer(&self, lock: BOOL) -> Result<()> {
        if lock.as_bool() {
            DLL_REF_COUNT.fetch_add(1, Ordering::SeqCst);
        } else {
            DLL_REF_COUNT.fetch_sub(1, Ordering::SeqCst);
        }
        Ok(())
    }
}

#[no_mangle]
pub extern "system" fn DllGetClassObject(
    rclsid: *const GUID,
    riid: *const GUID,
    ppv: *mut *mut c_void,
) -> HRESULT {
    unsafe {
        if *rclsid == CLSID_HASH_PROPERTY_HANDLER {
            let factory = Factory;
            let unknown: IUnknown = factory.into();
            unknown.query(&*riid, ppv)
        } else {
            CLASS_E_CLASSNOTAVAILABLE
        }
    }
}

#[no_mangle]
pub extern "system" fn DllCanUnloadNow() -> HRESULT {
    if DLL_REF_COUNT.load(Ordering::SeqCst) == 0 {
        S_OK
    } else {
        S_FALSE
    }
}

fn get_extensions_to_register() -> Vec<String> {
    let mut hkey = HKEY::default();
    let status = unsafe {
        RegOpenKeyExW(
            HKEY_CURRENT_USER,
            w!("Software\\WinExplorerHash"),
            0,
            KEY_READ,
            &mut hkey,
        )
    };

    if status.is_ok() {
        let mut data_size = 0u32;
        let mut reg_type = 0u32;
        
        let query_status = unsafe {
            RegQueryValueExW(
                hkey,
                w!("Extensions"),
                None,
                Some(&mut reg_type),
                None,
                Some(&mut data_size),
            )
        };

        if query_status.is_ok() && reg_type == REG_SZ.0 && data_size > 0 {
            let mut data = vec![0u8; data_size as usize];
            let query_status2 = unsafe {
                RegQueryValueExW(
                    hkey,
                    w!("Extensions"),
                    None,
                    Some(&mut reg_type),
                    Some(data.as_mut_ptr()),
                    Some(&mut data_size),
                )
            };

            unsafe { let _ = RegCloseKey(hkey); }

            if query_status2.is_ok() {
                let u16_slice = unsafe {
                    std::slice::from_raw_parts(data.as_ptr() as *const u16, data.len() / 2)
                };
                let len = u16_slice.iter().position(|&c| c == 0).unwrap_or(u16_slice.len());
                let s = String::from_utf16_lossy(&u16_slice[..len]);
                if !s.trim().is_empty() {
                    return s.split(',').map(|ext| {
                        let mut e = ext.trim().to_string();
                        if !e.starts_with('.') && !e.is_empty() {
                            e.insert(0, '.');
                        }
                        e
                    }).filter(|e| !e.is_empty()).collect();
                }
            }
        } else {
            unsafe { let _ = RegCloseKey(hkey); }
        }
    }

    vec![
        ".pdf".to_string(), ".png".to_string(), ".jpg".to_string(), ".jpeg".to_string(),
        ".txt".to_string(), ".csv".to_string(), ".xls".to_string(), ".xlsx".to_string()
    ]
}

#[no_mangle]
pub extern "system" fn DllRegisterServer() -> HRESULT {
    unsafe {
        let mut path_buf = [0u16; MAX_PATH as usize];
        GetModuleFileNameW(DLL_INSTANCE, &mut path_buf);

        // Find length
        let mut path_len = 0;
        while path_len < path_buf.len() && path_buf[path_len] != 0 {
            path_len += 1;
        }

        let path_bytes = std::slice::from_raw_parts(path_buf.as_ptr() as *const u8, path_len * 2);

        // 1. Register CLSID
        let mut hkey: HKEY = Default::default();
        if RegCreateKeyExW(
            HKEY_CLASSES_ROOT,
            w!("CLSID\\{8E97E8B8-5A24-4FDB-AA9D-9F319BE24B02}\\InProcServer32"),
            0,
            PCWSTR::null(),
            REG_OPTION_NON_VOLATILE,
            KEY_ALL_ACCESS,
            None,
            &mut hkey,
            None,
        ).is_err() {
            return E_FAIL;
        }

        let _ = RegSetValueExW(
            hkey,
            PCWSTR::null(),
            0,
            REG_SZ,
            Some(path_bytes),
        );

        let threading_model = "Apartment";
        let mut tm_bytes: Vec<u8> = threading_model.encode_utf16().flat_map(|c| c.to_ne_bytes()).collect();
        tm_bytes.push(0); tm_bytes.push(0);

        let _ = RegSetValueExW(
            hkey,
            w!("ThreadingModel"),
            0,
            REG_SZ,
            Some(&tm_bytes),
        );
        let _ = RegCloseKey(hkey);

        let extensions = get_extensions_to_register();

        for ext in extensions.iter() {
            // 2. Register Property Handler for specific extension
            let mut hkey_ext: HKEY = Default::default();
            let subkey_ext = format!("{}\\shellex\\PropertyHandler\0", ext);
            let subkey_ext_wide: Vec<u16> = subkey_ext.encode_utf16().collect();
            
            if RegCreateKeyExW(
                HKEY_CLASSES_ROOT,
                PCWSTR::from_raw(subkey_ext_wide.as_ptr()),
                0,
                None,
                REG_OPTION_NON_VOLATILE,
                KEY_ALL_ACCESS,
                None,
                &mut hkey_ext,
                None,
            ).is_ok()
            {
                let clsid_str = "{8E97E8B8-5A24-4FDB-AA9D-9F319BE24B02}\0";
                let clsid_bytes: Vec<u8> = clsid_str.encode_utf16().flat_map(|c| c.to_ne_bytes()).collect();
                let _ = RegSetValueExW(
                    hkey_ext,
                    None,
                    0,
                    REG_SZ,
                    Some(&clsid_bytes),
                );
                let _ = RegCloseKey(hkey_ext);
            }

            // Also register in PropertySystem PropertyHandlers
            let mut hkey_ps: HKEY = Default::default();
            let subkey_ps = format!("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\{}\0", ext);
            let subkey_ps_wide: Vec<u16> = subkey_ps.encode_utf16().collect();
            
            if RegCreateKeyExW(
                windows::Win32::System::Registry::HKEY_LOCAL_MACHINE,
                PCWSTR::from_raw(subkey_ps_wide.as_ptr()),
                0,
                None,
                REG_OPTION_NON_VOLATILE,
                KEY_ALL_ACCESS,
                None,
                &mut hkey_ps,
                None,
            ).is_ok()
            {
                let clsid_str = "{8E97E8B8-5A24-4FDB-AA9D-9F319BE24B02}\0";
                let clsid_bytes: Vec<u8> = clsid_str.encode_utf16().flat_map(|c| c.to_ne_bytes()).collect();
                let _ = RegSetValueExW(
                    hkey_ps,
                    None,
                    0,
                    REG_SZ,
                    Some(&clsid_bytes),
                );
                let _ = RegCloseKey(hkey_ps);
            }
        } // 3. Register Schema
        // Schema is typically located alongside the DLL. We replace "DLLname.dll" with "hash_schema.propdesc"
        // Since doing path manipulation in raw UTF-16 is tedious, we use String.
        let path_str = String::from_utf16_lossy(&path_buf[..path_len]);
        let dir = match std::path::Path::new(&path_str).parent() {
            Some(p) => p.to_string_lossy().to_string(),
            None => return E_FAIL,
        };
        let schema_path = format!("{}\\hash_schema.propdesc", dir);
        let hstr = windows::core::HSTRING::from(schema_path);

        if PSRegisterPropertySchema(PCWSTR::from_raw(hstr.as_ptr())).is_err() {
            // Non-fatal, perhaps schema is already registered
        }

        S_OK
    }
}

#[no_mangle]
pub extern "system" fn DllUnregisterServer() -> HRESULT {
    unsafe {
        let _ = RegDeleteTreeW(
            HKEY_CLASSES_ROOT,
            w!("CLSID\\{8E97E8B8-5A24-4FDB-AA9D-9F319BE24B02}"),
        );

        let extensions = get_extensions_to_register();

        for ext in extensions.iter() {
            let subkey_ext = format!("{}\\shellex\\PropertyHandler\0", ext);
            let subkey_ext_wide: Vec<u16> = subkey_ext.encode_utf16().collect();
            let _ = RegDeleteTreeW(
                HKEY_CLASSES_ROOT,
                PCWSTR::from_raw(subkey_ext_wide.as_ptr()),
            );

            let subkey_ps = format!("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\{}\0", ext);
            let subkey_ps_wide: Vec<u16> = subkey_ps.encode_utf16().collect();
            let _ = RegDeleteTreeW(
                windows::Win32::System::Registry::HKEY_LOCAL_MACHINE,
                PCWSTR::from_raw(subkey_ps_wide.as_ptr()),
            );
        }

        // Best effort schema path reconstruction
        let mut path_buf = [0u16; MAX_PATH as usize];
        if GetModuleFileNameW(DLL_INSTANCE, &mut path_buf) > 0 {
            let mut path_len = 0;
            while path_len < path_buf.len() && path_buf[path_len] != 0 {
                path_len += 1;
            }
            let path_str = String::from_utf16_lossy(&path_buf[..path_len]);
            if let Some(p) = std::path::Path::new(&path_str).parent() {
                let schema_path = format!("{}\\hash_schema.propdesc", p.display());
                let hstr = windows::core::HSTRING::from(schema_path);
                let _ = PSUnregisterPropertySchema(PCWSTR::from_raw(hstr.as_ptr()));
            }
        }
    }
    S_OK
}
