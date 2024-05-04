use anyhow::{anyhow, Context};
use libloading::{Library, Symbol};
use std::{env, thread};
use windows::core::GUID;
use windows::core::*;
use windows::Win32::System::Com::*;
use windows::Win32::System::Ole::*;

const LOCALE_USER_DEFAULT: u32 = 0x0400;
const LOCALE_SYSTEM_DEFAULT: u32 = 0x0800;

fn main() -> anyhow::Result<()> {
    unsafe {
        let res = CoInitialize(None);

        let _com = DeferCoUninitialize;

        let clsid = CLSIDFromProgID(PCWSTR::from_raw(
            HSTRING::from("PocketOutlook.Application").as_ptr(),
        ))
        .with_context(|| "CLSIDFromProgID")?;
        print!("clsid: {:?}\n", clsid);

        let guid = GUID::from_values(
            0x0006308B,
            0x0000,
            0x0000,
            [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
        );

        let outlook = CoCreateInstance(&guid, None, CLSCTX_LOCAL_SERVER)
            .with_context(|| "CoCreateInstance")?;
        let outlook = IDispatchWrapper(outlook);
        //let _outlook_quit = DeferOutlookQuit(&outlook);

        // Example operation: Creating and sending an email
        let mail_item = outlook.call("CreateItem", vec![0.into()])?; // 0 corresponds to MailItem
        let mail = mail_item.idispatch()?;

        mail.put("To", vec!["example@example.com".into()])?;
        mail.put("Subject", vec!["Test subject".into()])?;
        mail.put("Body", vec!["Hello from Rust!".into()])?;

        mail.call("Send", vec![])?;

        Ok(())
    }
}

// Helper structs and methods as before, adapted for Outlook if necessary
pub struct DeferCoUninitialize;

impl Drop for DeferCoUninitialize {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

pub struct IDispatchWrapper(pub IDispatch);

impl IDispatchWrapper {
    pub fn invoke(
        &self,
        flags: DISPATCH_FLAGS,
        name: &str,
        mut args: Vec<Variant>,
    ) -> anyhow::Result<Variant> {
        unsafe {
            let mut dispid = 0;
            self.0
                .GetIDsOfNames(
                    &GUID::default(),
                    &PCWSTR::from_raw(HSTRING::from(name).as_ptr()),
                    1,
                    LOCALE_USER_DEFAULT,
                    &mut dispid,
                )
                .with_context(|| "GetIDsOfNames")?;

            let mut dp = DISPPARAMS::default();
            let mut dispid_named = DISPID_PROPERTYPUT;

            if !args.is_empty() {
                args.reverse();
                dp.cArgs = args.len() as u32;
                dp.rgvarg = args.as_mut_ptr() as *mut VARIANT;

                // Handle special-case for property-puts!
                if (flags & DISPATCH_PROPERTYPUT) != DISPATCH_FLAGS(0) {
                    dp.cNamedArgs = 1;
                    dp.rgdispidNamedArgs = &mut dispid_named;
                }
            }

            let mut result = VARIANT::default();
            self.0
                .Invoke(
                    dispid,
                    &GUID::default(),
                    LOCALE_SYSTEM_DEFAULT,
                    flags,
                    &dp,
                    Some(&mut result),
                    None,
                    None,
                )
                .with_context(|| "Invoke")?;

            Ok(Variant(result))
        }
    }

    pub fn get(&self, name: &str) -> anyhow::Result<Variant> {
        self.invoke(DISPATCH_PROPERTYGET, name, vec![])
    }

    pub fn int(&self, name: &str) -> anyhow::Result<i32> {
        let result = self.get(name)?;
        result.int()
    }

    pub fn bool(&self, name: &str) -> anyhow::Result<bool> {
        let result = self.get(name)?;
        result.bool()
    }

    pub fn string(&self, name: &str) -> anyhow::Result<String> {
        let result = self.get(name)?;
        result.string()
    }

    pub fn put(&self, name: &str, args: Vec<Variant>) -> anyhow::Result<Variant> {
        self.invoke(DISPATCH_PROPERTYPUT, name, args)
    }

    pub fn call(&self, name: &str, args: Vec<Variant>) -> anyhow::Result<Variant> {
        self.invoke(DISPATCH_METHOD, name, args)
    }
}

pub struct Variant(VARIANT);

impl From<bool> for Variant {
    fn from(value: bool) -> Self {
        Self(value.into())
    }
}

impl From<i32> for Variant {
    fn from(value: i32) -> Self {
        Self(value.into())
    }
}

impl From<&str> for Variant {
    fn from(value: &str) -> Self {
        Self(BSTR::from(value).into())
    }
}

impl From<&String> for Variant {
    fn from(value: &String) -> Self {
        Self(BSTR::from(value).into())
    }
}

impl Variant {
    pub fn bool(&self) -> anyhow::Result<bool> {
        Ok(bool::try_from(&self.0)?)
    }

    pub fn int(&self) -> anyhow::Result<i32> {
        Ok(i32::try_from(&self.0)?)
    }

    pub fn string(&self) -> anyhow::Result<String> {
        Ok(BSTR::try_from(&self.0)?.to_string())
    }

    pub fn idispatch(&self) -> anyhow::Result<IDispatchWrapper> {
        Ok(IDispatchWrapper(IDispatch::try_from(&self.0)?))
    }

    pub fn vt(&self) -> u16 {
        unsafe { self.0.as_raw().Anonymous.Anonymous.vt }
    }
}
