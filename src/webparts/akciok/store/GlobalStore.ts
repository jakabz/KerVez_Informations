import * as React from 'react';
import { createStore } from 'zustand/vanilla';
import { useStore } from 'zustand';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import pnpListsServices from '../services/pnpListServices';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { IShopListItem } from '../models/IShop';

export type PanelMode = 'create' | 'edit' | 'view';
export type ViewMode = 'dashboard' | 'full';
export type ActionsUiViewMode = 'cards' | 'list';
export type ActionSortBy =
    'actionTypeAsc'
    | 'actionTypeDesc'
    | 'endDateDesc'
    | 'endDateAsc'
    | 'responsiblePersonAsc'
    | 'responsiblePersonDesc';

type GlobalState = {
    context: WebPartContext;
    actionsListId: string;
    shopsListId: string;
    goodsListId: string;
    shopListItems: IShopListItem[];
    shopTypeChoices: string[];
    relatedDocumentLibraryId: string;
    tasksListPage: string;
    editorGroup: IPropertyFieldGroupOrPerson[];
    isEditor: boolean;
    isPanelOpen: boolean;
    panelMode: PanelMode;
    selectedActionId: number | undefined;
    refreshKey: number;
    actionSortBy: ActionSortBy;
    hideExpiredActions: boolean;
    hideArchivedActions: boolean;
    selectedActionTypes: string[];
    availableActionTypes: string[];
    actionSearchText: string;
    viewMode: ViewMode;
    actionsUiViewMode: ActionsUiViewMode;
    setContext: (context: WebPartContext) => void;
    setActionsListId: (actionsListId: string) => void;
    setShopsListId: (shopsListId: string) => void;
    setGoodsListId: (goodsListId: string) => void;
    setRelatedDocumentLibraryId: (relatedDocumentLibraryId: string) => void;
    setTasksListPage: (tasksListPage: string) => void;
    setEditorGroup: (editorGroup: IPropertyFieldGroupOrPerson[]) => void;
    setIsEditor: (isEditor: boolean) => void;
    setPanelMode: (panelMode: PanelMode) => void;
    openPanel: () => void;
    openCreatePanel: () => void;
    openActionPanel: (itemId: number, mode: Exclude<PanelMode, 'create'>) => void;
    closePanel: () => void;
    triggerRefresh: () => void;
    setActionSortBy: (actionSortBy: ActionSortBy) => void;
    setHideExpiredActions: (hideExpiredActions: boolean) => void;
    setHideArchivedActions: (hideArchivedActions: boolean) => void;
    setSelectedActionTypes: (selectedActionTypes: string[]) => void;
    setAvailableActionTypes: (availableActionTypes: string[]) => void;
    setActionSearchText: (actionSearchText: string) => void;
    setViewMode: (viewMode: ViewMode) => void;
    setActionsUiViewMode: (actionsUiViewMode: ActionsUiViewMode) => void;
    resetState: () => void;
}

type GlobalStoreApi = ReturnType<typeof createGlobalStore>;

const createInitialState = (): Omit<GlobalState, 'setContext' | 'setActionsListId' | 'setShopsListId' | 'setGoodsListId' | 'setRelatedDocumentLibraryId' | 'setTasksListPage' | 'setEditorGroup' | 'setIsEditor' | 'setPanelMode' | 'openPanel' | 'openCreatePanel' | 'openActionPanel' | 'closePanel' | 'triggerRefresh' | 'setActionSortBy' | 'setHideExpiredActions' | 'setHideArchivedActions' | 'setSelectedActionTypes' | 'setAvailableActionTypes' | 'setActionSearchText' | 'setViewMode' | 'setActionsUiViewMode' | 'resetState'> => ({
    context: {} as WebPartContext,
    actionsListId: '',
    shopsListId: '',
    goodsListId: '',
    relatedDocumentLibraryId: '',
    tasksListPage: '',
    shopListItems: [],
    shopTypeChoices: [],
    editorGroup: [],
    isEditor: false,
    isPanelOpen: false,
    panelMode: 'create',
    selectedActionId: undefined,
    refreshKey: 0,
    actionSortBy: 'endDateDesc',
    hideExpiredActions: true,
    hideArchivedActions: true,
    selectedActionTypes: [],
    availableActionTypes: [],
    actionSearchText: '',
    viewMode: 'full',
    actionsUiViewMode: 'cards',
});

export const createGlobalStore = () => createStore<GlobalState>((set, get) => ({
    ...createInitialState(),
    setContext: (context: WebPartContext) => {
        set({ context });
    },
    setActionsListId: (actionsListId: string) => {
        set({ actionsListId });
    },
    setShopsListId: async (shopsListId: string) => {
        set({ shopsListId });

        if (!shopsListId) {
            set({ shopListItems: [], shopTypeChoices: [] });
            return;
        }

        try {
            const [shopListItems, shopTypeChoices] = await Promise.all([
                pnpListsServices.getShops(shopsListId),
                pnpListsServices.getShopTypeChoices(shopsListId),
            ]);
            set({ shopListItems, shopTypeChoices });
        } catch {
            set({ shopListItems: [], shopTypeChoices: [] });
        }
    },
    setGoodsListId: (goodsListId: string) => {
        set({ goodsListId });
    },
    setRelatedDocumentLibraryId: (relatedDocumentLibraryId: string) => {
        set({ relatedDocumentLibraryId });
    },
    setTasksListPage: (tasksListPage: string) => {
        set({ tasksListPage });
    },
    setEditorGroup: async (editorGroup: IPropertyFieldGroupOrPerson[]) => {
        set({ editorGroup });
        if(editorGroup?.length > 0) {
            const currentUser = get().context.pageContext.user;
            const isEditor = await pnpListsServices.isMemberOfGroups(editorGroup, currentUser.loginName);
            set({ isEditor });
        } else {
            set({ isEditor: false });
        }
    },
    setIsEditor: (isEditor: boolean) => {
        set({ isEditor });
    },
    setPanelMode: (panelMode: PanelMode) => {
        set({ panelMode });
    },
    openPanel: () => set({ isPanelOpen: true, panelMode: 'create', selectedActionId: undefined }),
    openCreatePanel: () => set({ isPanelOpen: true, panelMode: 'create', selectedActionId: undefined }),
    openActionPanel: (itemId: number, mode: Exclude<PanelMode, 'create'>) => set({ isPanelOpen: true, panelMode: mode, selectedActionId: itemId }),
    closePanel: () => set({ isPanelOpen: false, panelMode: 'create', selectedActionId: undefined }),
    triggerRefresh: () => set(s => ({ refreshKey: s.refreshKey + 1 })),
    setActionSortBy: (actionSortBy: ActionSortBy) => set({ actionSortBy }),
    setHideExpiredActions: (hideExpiredActions: boolean) => set({ hideExpiredActions }),
    setHideArchivedActions: (hideArchivedActions: boolean) => set({ hideArchivedActions }),
    setSelectedActionTypes: (selectedActionTypes: string[]) => set({ selectedActionTypes }),
    setAvailableActionTypes: (availableActionTypes: string[]) => set({ availableActionTypes }),
    setActionSearchText: (actionSearchText: string) => set({ actionSearchText }),
    setViewMode: (viewMode: ViewMode) => set({ viewMode }),
    setActionsUiViewMode: (actionsUiViewMode: ActionsUiViewMode) => set({ actionsUiViewMode }),
    resetState: () => set(createInitialState()),
}));

const GlobalStoreContext = React.createContext<GlobalStoreApi | null>(null);

export const GlobalStoreProvider: React.FunctionComponent<React.PropsWithChildren<{ store: GlobalStoreApi }>> = ({ store, children }) => (
    React.createElement(GlobalStoreContext.Provider, { value: store }, children)
);

export function useGlobalState(): GlobalState;
export function useGlobalState<T>(selector: (state: GlobalState) => T): T;
export function useGlobalState<T>(selector?: (state: GlobalState) => T): T | GlobalState {
    const store = React.useContext(GlobalStoreContext);

    if (!store) {
        throw new Error('useGlobalState must be used within GlobalStoreProvider.');
    }

    if (selector) {
        return useStore(store, selector);
    }

    return useStore(store);
}
