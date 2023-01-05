// This is undocumented - but needed to get client url and mode
// so we must know the mode as most of the methods are not implemented in Authoring mode
export interface ContextEx {
    page: {
        getClientUrl: () => string;
        entityId: string;
        entityTypeName: string;
    };
    mode: {
        isAuthoringMode: boolean;
    }
}