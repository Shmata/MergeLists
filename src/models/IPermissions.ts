export interface IBasePermissions {
    Low: number;
    High: number;
}
export enum PermissionKind {
    /**
     * Has no permissions on the Site. Not available through the user interface.
     */
    EmptyMask = 0,
    /**
     * View items in lists, documents in document libraries, and Web discussion comments.
     */
    ViewListItems = 1,
    /**
     * Add items to lists, documents to document libraries, and Web discussion comments.
     */
    AddListItems = 2,
    /**
     * Edit items in lists, edit documents in document libraries, edit Web discussion comments
     * in documents, and customize Web Part Pages in document libraries.
     */
    EditListItems = 3,
    /**
     * Delete items from a list, documents from a document library, and Web discussion
     * comments in documents.
     */
    DeleteListItems = 4,
    /**
     * Approve a minor version of a list item or document.
     */
    ApproveItems = 5,
    /**
     * View the source of documents with server-side file handlers.
     */
    OpenItems = 6,
    /**
     * View past versions of a list item or document.
     */
    ViewVersions = 7,
    /**
     * Delete past versions of a list item or document.
     */
    DeleteVersions = 8,
    /**
     * Discard or check in a document which is checked out to another user.
     */
    CancelCheckout = 9,
    /**
     * Create, change, and delete personal views of lists.
     */
    ManagePersonalViews = 10,
    /**
     * Create and delete lists, add or remove columns in a list, and add or remove public views of a list.
     */
    ManageLists = 12,
    /**
     * View forms, views, and application pages, and enumerate lists.
     */
    ViewFormPages = 13,
    /**
     * Make content of a list or document library retrieveable for anonymous users through SharePoint search.
     * The list permissions in the site do not change.
     */
    AnonymousSearchAccessList = 14,
    /**
     * Allow users to open a Site, list, or folder to access items inside that container.
     */
    Open = 17,
    /**
     * View pages in a Site.
     */
    ViewPages = 18,
    /**
     * Add, change, or delete HTML pages or Web Part Pages, and edit the Site using
     * a Windows SharePoint Services compatible editor.
     */
    AddAndCustomizePages = 19,
    /**
     * Apply a theme or borders to the entire Site.
     */
    ApplyThemeAndBorder = 20,
    /**
     * Apply a style sheet (.css file) to the Site.
     */
    ApplyStyleSheets = 21,
    /**
     * View reports on Site usage.
     */
    ViewUsageData = 22,
    /**
     * Create a Site using Self-Service Site Creation.
     */
    CreateSSCSite = 23,
    /**
     * Create subsites such as team sites, Meeting Workspace sites, and Document Workspace sites.
     */
    ManageSubwebs = 24,
    /**
     * Create a group of users that can be used anywhere within the site collection.
     */
    CreateGroups = 25,
    /**
     * Create and change permission levels on the Site and assign permissions to users
     * and groups.
     */
    ManagePermissions = 26,
    /**
     * Enumerate files and folders in a Site using Microsoft Office SharePoint Designer
     * and WebDAV interfaces.
     */
    BrowseDirectories = 27,
    /**
     * View information about users of the Site.
     */
    BrowseUserInfo = 28,
    /**
     * Add or remove personal Web Parts on a Web Part Page.
     */
    AddDelPrivateWebParts = 29,
    /**
     * Update Web Parts to display personalized information.
     */
    UpdatePersonalWebParts = 30,
    /**
     * Grant the ability to perform all administration tasks for the Site as well as
     * manage content, activate, deactivate, or edit properties of Site scoped Features
     * through the object model or through the user interface (UI). When granted on the
     * root Site of a Site Collection, activate, deactivate, or edit properties of
     * site collection scoped Features through the object model. To browse to the Site
     * Collection Features page and activate or deactivate Site Collection scoped Features
     * through the UI, you must be a Site Collection administrator.
     */
    ManageWeb = 31,
    /**
     * Content of lists and document libraries in the Web site will be retrieveable for anonymous users through
     * SharePoint search if the list or document library has AnonymousSearchAccessList set.
     */
    AnonymousSearchAccessWebLists = 32,
    /**
     * Use features that launch client applications. Otherwise, users must work on documents
     * locally and upload changes.
     */
    UseClientIntegration = 37,
    /**
     * Use SOAP, WebDAV, or Microsoft Office SharePoint Designer interfaces to access the Site.
     */
    UseRemoteAPIs = 38,
    /**
     * Manage alerts for all users of the Site.
     */
    ManageAlerts = 39,
    /**
     * Create e-mail alerts.
     */
    CreateAlerts = 40,
    /**
     * Allows a user to change his or her user information, such as adding a picture.
     */
    EditMyUserInfo = 41,
    /**
     * Enumerate permissions on Site, list, folder, document, or list item.
     */
    EnumeratePermissions = 63,
    /**
     * Has all permissions on the Site. Not available through the user interface.
     */
    FullMask = 65
}

export function hasPermissions(value: IBasePermissions, perm: PermissionKind) {
    if (!perm) {
        return true;
    }
    if (perm === PermissionKind.FullMask) {
        return (value.High & 32767) === 32767 && value.Low === 65535;
    }
    perm = perm - 1;
    var num = 1;
    if (perm >= 0 && perm < 32) {
        num = num << perm;
        return 0 !== (value.Low & num);
    }
    else if (perm >= 32 && perm < 64) {
        num = num << perm - 32;
        return 0 !== (value.High & num);
    }
    return false;
}