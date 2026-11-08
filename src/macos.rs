//! Native macOS pieces that WKWebView does not provide on its own:
//! the main menu (which is what makes Cut, Copy, Paste, and Undo reach the web view),
//! the Save panel, and printing.

use std::cell::RefCell;
use std::path::PathBuf;
use std::ptr::null_mut;

use objc2::rc::Retained;
use objc2::runtime::{AnyObject, NSObject, Sel};
use objc2::{declare_class, msg_send_id, mutability, sel, ClassType, DeclaredClass};
use objc2_app_kit::{
    NSApplication, NSEventModifierFlags, NSMenu, NSMenuItem, NSModalResponseOK, NSPrintInfo,
    NSSavePanel,
};
use objc2_foundation::{MainThreadMarker, NSString};
use wry::{WebView, WebViewExtMacOS};

const PAGE_MARGIN_POINTS: f64 = 72.0; // 1in, matching the on screen document page

/// Asks where to save a file. `Ok(None)` means the person cancelled.
pub fn prompt_save_path(default_name: &str) -> Result<Option<PathBuf>, String> {
    let mtm = MainThreadMarker::new().ok_or("The Save panel must run on the main thread")?;

    unsafe {
        let panel = NSSavePanel::savePanel(mtm);
        panel.setNameFieldStringValue(&NSString::from_str(default_name));
        panel.setCanCreateDirectories(true);
        panel.setExtensionHidden(false);

        if panel.runModal() != NSModalResponseOK {
            return Ok(None);
        }

        let path = panel
            .URL()
            .and_then(|url| url.path())
            .ok_or("The Save panel did not return a file location")?;
        Ok(Some(PathBuf::from(path.to_string())))
    }
}

/// Opens the system print panel, which also offers Save as PDF.
pub fn print_webview(webview: &WebView, title: &str) {
    unsafe {
        let wk_webview = webview.webview();
        let window = webview.ns_window();

        let print_info = NSPrintInfo::sharedPrintInfo();
        print_info.setTopMargin(PAGE_MARGIN_POINTS);
        print_info.setRightMargin(PAGE_MARGIN_POINTS);
        print_info.setBottomMargin(PAGE_MARGIN_POINTS);
        print_info.setLeftMargin(PAGE_MARGIN_POINTS);
        print_info.setVerticallyCentered(false);

        let operation = wk_webview.printOperationWithPrintInfo(&print_info);
        operation.setJobTitle(Some(&NSString::from_str(title))); // Default Save as PDF name
        operation.setShowsPrintPanel(true);
        operation.setShowsProgressPanel(true);

        // The print view starts with an empty frame, and WebKit then prints blank pages.
        if let Some(view) = operation.view() {
            view.setFrame(wk_webview.frame());
        }

        operation.runOperationModalForWindow_delegate_didRunSelector_contextInfo(
            &window,
            None,
            None,
            null_mut(),
        );
    }
}

thread_local! {
    static FILE_COMMAND: RefCell<Option<Box<dyn Fn(&'static str)>>> = RefCell::new(None);
}

fn run_file_command(command: &'static str) {
    FILE_COMMAND.with(|handler| {
        if let Some(handler) = handler.borrow().as_ref() {
            handler(command);
        }
    });
}

declare_class!(
    /// Receives the File menu commands, which the page carries out.
    struct FileMenuTarget;

    // SAFETY: NSObject has no subclassing requirements, and this class has no Drop.
    unsafe impl ClassType for FileMenuTarget {
        type Super = NSObject;
        type Mutability = mutability::MainThreadOnly;
        const NAME: &'static str = "OfficeLiteFileMenuTarget";
    }

    impl DeclaredClass for FileMenuTarget {}

    unsafe impl FileMenuTarget {
        #[method(officeLiteNew:)]
        fn new_file(&self, _sender: Option<&AnyObject>) {
            run_file_command("new");
        }

        #[method(officeLiteOpen:)]
        fn open_file(&self, _sender: Option<&AnyObject>) {
            run_file_command("open");
        }

        #[method(officeLiteSave:)]
        fn save_file(&self, _sender: Option<&AnyObject>) {
            run_file_command("save");
        }

        #[method(officeLitePrint:)]
        fn print_file(&self, _sender: Option<&AnyObject>) {
            run_file_command("print");
        }
    }
);

/// Installs the menu bar. Edit and Window commands use a nil target, so they travel the
/// responder chain to whatever has focus, including the web view. File commands go to
/// `on_file_command`.
pub fn install_main_menu(on_file_command: impl Fn(&'static str) + 'static) {
    let Some(mtm) = MainThreadMarker::new() else {
        return;
    };
    FILE_COMMAND.with(|handler| *handler.borrow_mut() = Some(Box::new(on_file_command)));

    // Menu items do not retain their target, so this lives for the rest of the app's life.
    let file_target: &'static FileMenuTarget = unsafe {
        let this = mtm.alloc::<FileMenuTarget>().set_ivars(());
        let target: Retained<FileMenuTarget> = msg_send_id![super(this), init];
        &*Retained::into_raw(target)
    };
    let file_target_object: &AnyObject = file_target;

    const COMMAND: NSEventModifierFlags = NSEventModifierFlags::NSEventModifierFlagCommand;
    const SHIFT: NSEventModifierFlags = NSEventModifierFlags::NSEventModifierFlagShift;
    const OPTION: NSEventModifierFlags = NSEventModifierFlags::NSEventModifierFlagOption;
    const CONTROL: NSEventModifierFlags = NSEventModifierFlags::NSEventModifierFlagControl;

    let app_menu = submenu(
        mtm,
        "OfficeLite",
        &[
            Entry::Item("About OfficeLite", sel!(orderFrontStandardAboutPanel:), "", COMMAND),
            Entry::Separator,
            Entry::Item("Hide OfficeLite", sel!(hide:), "h", COMMAND),
            Entry::Item("Hide Others", sel!(hideOtherApplications:), "h", combine(COMMAND, OPTION)),
            Entry::Item("Show All", sel!(unhideAllApplications:), "", COMMAND),
            Entry::Separator,
            Entry::Item("Quit OfficeLite", sel!(terminate:), "q", COMMAND),
        ],
    );

    let file_menu = submenu_with_target(
        mtm,
        "File",
        &[
            Entry::Item("New", sel!(officeLiteNew:), "n", COMMAND),
            Entry::Item("Open…", sel!(officeLiteOpen:), "o", COMMAND),
            Entry::Separator,
            Entry::Item("Save…", sel!(officeLiteSave:), "s", COMMAND),
            Entry::Separator,
            Entry::Item("Print…", sel!(officeLitePrint:), "p", COMMAND),
        ],
        Some(file_target_object),
    );

    let edit_menu = submenu(
        mtm,
        "Edit",
        &[
            Entry::Item("Undo", sel!(undo:), "z", COMMAND),
            Entry::Item("Redo", sel!(redo:), "z", combine(COMMAND, SHIFT)),
            Entry::Separator,
            Entry::Item("Cut", sel!(cut:), "x", COMMAND),
            Entry::Item("Copy", sel!(copy:), "c", COMMAND),
            Entry::Item("Paste", sel!(paste:), "v", COMMAND),
            Entry::Item(
                "Paste and Match Style",
                sel!(pasteAsPlainText:),
                "v",
                combine(combine(COMMAND, SHIFT), OPTION),
            ),
            Entry::Item("Delete", sel!(delete:), "", COMMAND),
            Entry::Item("Select All", sel!(selectAll:), "a", COMMAND),
        ],
    );

    let window_menu = submenu(
        mtm,
        "Window",
        &[
            Entry::Item("Minimize", sel!(performMiniaturize:), "m", COMMAND),
            Entry::Item("Zoom", sel!(performZoom:), "", COMMAND),
            Entry::Item("Close", sel!(performClose:), "w", COMMAND),
            Entry::Separator,
            Entry::Item("Enter Full Screen", sel!(toggleFullScreen:), "f", combine(COMMAND, CONTROL)),
        ],
    );

    unsafe {
        let menu_bar = NSMenu::initWithTitle(mtm.alloc(), &NSString::from_str(""));
        for menu in [&app_menu, &file_menu, &edit_menu, &window_menu] {
            let holder = NSMenuItem::new(mtm);
            holder.setSubmenu(Some(menu));
            menu_bar.addItem(&holder);
        }

        let app = NSApplication::sharedApplication(mtm);
        app.setMainMenu(Some(&menu_bar));
        app.setWindowsMenu(Some(&window_menu));
    }
}

enum Entry {
    Item(&'static str, Sel, &'static str, NSEventModifierFlags),
    Separator,
}

const fn combine(a: NSEventModifierFlags, b: NSEventModifierFlags) -> NSEventModifierFlags {
    NSEventModifierFlags(a.0 | b.0)
}

fn submenu(mtm: MainThreadMarker, title: &str, entries: &[Entry]) -> Retained<NSMenu> {
    submenu_with_target(mtm, title, entries, None)
}

fn submenu_with_target(
    mtm: MainThreadMarker,
    title: &str,
    entries: &[Entry],
    target: Option<&AnyObject>,
) -> Retained<NSMenu> {
    unsafe {
        let menu = NSMenu::initWithTitle(mtm.alloc(), &NSString::from_str(title));
        for entry in entries {
            match entry {
                Entry::Separator => menu.addItem(&NSMenuItem::separatorItem(mtm)),
                Entry::Item(label, action, key, modifiers) => {
                    let item = NSMenuItem::initWithTitle_action_keyEquivalent(
                        mtm.alloc(),
                        &NSString::from_str(label),
                        Some(*action),
                        &NSString::from_str(key),
                    );
                    item.setKeyEquivalentModifierMask(*modifiers);
                    item.setTarget(target);
                    menu.addItem(&item);
                }
            }
        }
        menu
    }
}
