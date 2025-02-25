export interface DriveItemResponse {
    "@odata.context": string;
    value: DriveItem[];
}

export interface DriveItem {
    createdDateTime: string;
    id: string;
    lastModifiedDateTime: string;
    name: string;
    webUrl: string;
    size: number;
    createdBy: {
        user: {
            email: string;
            displayName: string;
        };
    };
    lastModifiedBy: {
        user: {
            email: string;
            displayName: string;
        };
    };
    parentReference: {
        driveType: string;
        driveId: string;
        id: string;
        siteId: string;
    };
    file: {
        mimeType: string;
        hashes: Record<string, string>;
    };
    fileSystemInfo: {
        createdDateTime: string;
        lastModifiedDateTime: string;
    };
    searchResult: Record<string, unknown>;
}

export interface CustomDriveItemResponse extends DriveItem {
    thumbnail: string;
}
