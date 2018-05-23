function Metadata(is_table, has_headers, rows, cols, has_template, has_data, excelType) {
    this.is_table = is_table;
    this.has_headers = has_headers;
    this.rows = rows;
    this.cols = cols;
    this.has_template = has_template;
    this.has_data = has_data;
    this.excelType = excelType;
}

function UsersPermissionWs(mail, permissionType) {
    this.mail = mail;
    this.permissionType = permissionType;
}

function FileList(filename, sequenceNumber, type, data) {
    this.filename = filename;
    this.sequenceNumber = sequenceNumber;
    this.type = type;
    this.data = data;
}

function Data(Encrypt, symmetricKeys, data) {
    this.Encrypt = Encrypt;
    this.symmetricKeys = symmetricKeys;
    this.data = data;
}

function SymmetricKeys(userId, Key) {
    this.userId = userId;
    this.Key = Key;
}

function PushRequest(id, description, owner, usersPermission, viewServer, fromAddin, optionalFileList, fileList, clientType, metadata) {
    this.id = id;
    this.description = description;
    this.owner = owner;
    this.usersPermission = usersPermission;
    this.viewServer = viewServer;
    this.fromAddin = fromAddin;
    this.optionalFileList = optionalFileList;
    this.fileList = fileList;
    this.clientType = clientType;
    this.metadata = metadata;
}

function PullRequest(id, viewServer, optionalFilenameList, clientSeqNum) {
    this.id = id;
    this.viewServer = viewServer;
    this.optionalFilenameList = optionalFilenameList;
    this.clientSeqNum = clientSeqNum;
}

function ObjEventCreation(id, description, owner, usersPermission, viewServer, fromAddin, clientType, metadata, encrypted) {
    this.id = id;
    this.description = description;
    this.owner = owner;
    this.usersPermission = usersPermission;
    this.viewServer = viewServer;
    this.fromAddin = fromAddin;
    this.clientType = clientType;
    this.metadata = metadata;
    this.encrypted = encrypted;
}

function ObjEventUpdate(viewId, viewServer, encrypted) {
    this.viewId = viewId;
    this.viewServer = viewServer;
    this.encrypted = encrypted;
}

function ViewObj(id, view_server) {
    this.id = id;
    this.view_server = view_server;
}