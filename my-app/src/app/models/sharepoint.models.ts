export interface SpList {
    displayName: string;
    id: string;
    name: string;
}

export interface SpDrive {
    displayName: string;
    id: string;
    name: string;
}

export interface SpConfig {
    drives: SpDrive[];
    lists: SpList[];
    tenantId: string;
    clientId: string;
    redirectUri: string;
    siteHostName: string;
    sitePath: string;
    siteId: string;
}