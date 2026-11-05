import * as React from 'react';
import {
    Callout,
    Checkbox,
    DefaultButton,
    DirectionalHint,
    Dropdown,
    IContextualMenuProps,
    IDropdownOption,
    IDropdownStyles,
    PrimaryButton,
    SearchBox,
    Stack,
    Text,
} from '@fluentui/react';
import styles from './Akciok.module.scss';
import { ActionsUiViewMode, useGlobalState } from '../store/GlobalStore';
export interface IAkciokListHeaderProps { }

const roundedDropdownStyles: Partial<IDropdownStyles> = {
    title: { borderRadius: 10 },
};

export const AkciokListHeader: React.FunctionComponent<IAkciokListHeaderProps> = (props: React.PropsWithChildren<IAkciokListHeaderProps>) => {

    const [isFilterCalloutOpen, setIsFilterCalloutOpen] = React.useState(false);
    const [isViewCalloutOpen, setIsViewCalloutOpen] = React.useState(false);
    const filterButtonWrapperRef = React.useRef<HTMLDivElement | null>(null);
    const viewButtonWrapperRef = React.useRef<HTMLDivElement | null>(null);

    const {
        isEditor,
        openPanel,
        actionsUiViewMode,
        setActionsUiViewMode,
        actionSortBy,
        setActionSortBy,
        hideExpiredActions,
        setHideExpiredActions,
        hideArchivedActions,
        setHideArchivedActions,
        selectedActionTypes,
        setSelectedActionTypes,
        availableActionTypes,
        actionSearchText,
        setActionSearchText,
    } = useGlobalState();

    const actionTypeOptions = React.useMemo<IDropdownOption[]>(() => (
        availableActionTypes.map((actionType) => ({ key: actionType, text: actionType }))
    ), [availableActionTypes]);

    const viewModeOptions: Array<{ key: ActionsUiViewMode; text: string; iconName: string }> = [
        { key: 'cards', text: 'Kártya', iconName: 'PaymentCard' },
        { key: 'list', text: 'Lista', iconName: 'BulletedList' },
    ];

    const currentViewLabel = React.useMemo(() => {
        const selected = viewModeOptions.find((option) => option.key === actionsUiViewMode);
        return selected ? selected.text : 'Nézet';
    }, [actionsUiViewMode]);

    const sortMenuProps: IContextualMenuProps = {
        items: [
            {
                key: 'actionTypeAsc',
                text: 'Információ típusa (A-Z)',
                canCheck: true,
                checked: actionSortBy === 'actionTypeAsc',
                onClick: () => setActionSortBy('actionTypeAsc'),
            },
            {
                key: 'actionTypeDesc',
                text: 'Információ típusa (Z-A)',
                canCheck: true,
                checked: actionSortBy === 'actionTypeDesc',
                onClick: () => setActionSortBy('actionTypeDesc'),
            },
            {
                key: 'endDateDesc',
                text: 'Lejárat dátuma (legkésőbbi elöl)',
                canCheck: true,
                checked: actionSortBy === 'endDateDesc',
                onClick: () => setActionSortBy('endDateDesc'),
            },
            {
                key: 'endDateAsc',
                text: 'Lejárat dátuma (legkorábbi elöl)',
                canCheck: true,
                checked: actionSortBy === 'endDateAsc',
                onClick: () => setActionSortBy('endDateAsc'),
            },
            {
                key: 'responsiblePersonAsc',
                text: 'Kapcsolattartó (A-Z)',
                canCheck: true,
                checked: actionSortBy === 'responsiblePersonAsc',
                onClick: () => setActionSortBy('responsiblePersonAsc'),
            },
            {
                key: 'responsiblePersonDesc',
                text: 'Kapcsolattartó (Z-A)',
                canCheck: true,
                checked: actionSortBy === 'responsiblePersonDesc',
                onClick: () => setActionSortBy('responsiblePersonDesc'),
            },
        ],
    };

    const viewMenuProps: IContextualMenuProps = {
        items: viewModeOptions.map((option) => ({
            key: option.key,
            text: option.text,
            //iconProps: { iconName: option.iconName },
            canCheck: true,
            checked: actionsUiViewMode === option.key,
            onClick: () => {
                setActionsUiViewMode(option.key);
                setIsViewCalloutOpen(false);
            },
        })),
    };

    return (
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" className={styles.akciokListHeader}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                <div ref={filterButtonWrapperRef}>
                    <DefaultButton
                        text="Szűrés"
                        iconProps={{ iconName: 'Filter' }}
                        onClick={() => {
                            setIsFilterCalloutOpen((isOpen) => !isOpen);
                            setIsViewCalloutOpen(false);
                        }}
                    />
                </div>
                <DefaultButton
                    text="Rendezés"
                    iconProps={{ iconName: 'Sort' }}
                    menuProps={sortMenuProps}
                />
                <div ref={viewButtonWrapperRef}>
                    <DefaultButton
                        text={currentViewLabel}
                        iconProps={{ iconName: 'View' }}
                        menuProps={viewMenuProps}
                    />
                </div>
                <SearchBox
                    placeholder="Keresés EAN kódra vagy információ címre"
                    value={actionSearchText}
                    onChange={(_, newValue) => setActionSearchText(newValue ?? '')}
                    onClear={() => setActionSearchText('')}
                    styles={{ root: { width: 320 } }}
                />
            </Stack>
            {isFilterCalloutOpen && filterButtonWrapperRef.current && (
                <Callout
                    target={filterButtonWrapperRef.current}
                    onDismiss={() => setIsFilterCalloutOpen(false)}
                    setInitialFocus
                    directionalHint={DirectionalHint.bottomLeftEdge}
                    gapSpace={8}
                >
                    <Stack tokens={{ childrenGap: 10 }} style={{ minWidth: 320, padding: 12 }}>
                        <Text variant="mediumPlus">Szűrési feltételek</Text>
                        <Checkbox
                            label="Lejárt információk elrejtése"
                            checked={hideExpiredActions}
                            onChange={(_, checked) => setHideExpiredActions(!!checked)}
                        />
                        <Checkbox
                            label="Archive információk elrejtése"
                            checked={hideArchivedActions}
                            onChange={(_, checked) => setHideArchivedActions(!!checked)}
                        />
                        <Dropdown
                            label="Információ típusa"
                            placeholder={actionTypeOptions.length ? 'Válassz típust' : 'Nincs elérhető típus'}
                            options={actionTypeOptions}
                            multiSelect
                            selectedKeys={selectedActionTypes}
                            disabled={!actionTypeOptions.length}
                            styles={roundedDropdownStyles}
                            onChange={(_, option) => {
                                if (!option) {
                                    return;
                                }

                                const optionKey = String(option.key);
                                const nextSelectedActionTypes = option.selected
                                    ? [...selectedActionTypes, optionKey]
                                    : selectedActionTypes.filter((value) => value !== optionKey);
                                setSelectedActionTypes(nextSelectedActionTypes);
                            }}
                        />
                    </Stack>
                </Callout>
            )}
            {isViewCalloutOpen && viewButtonWrapperRef.current && (
                <Callout
                    target={viewButtonWrapperRef.current}
                    onDismiss={() => setIsViewCalloutOpen(false)}
                    setInitialFocus
                    directionalHint={DirectionalHint.bottomLeftEdge}
                    gapSpace={8}
                >
                    <Stack tokens={{ childrenGap: 10 }} style={{ minWidth: 280, padding: 12 }}>
                        <Text variant="mediumPlus">Nézet kiválasztása</Text>
                        <Stack tokens={{ childrenGap: 8 }}>
                            {viewModeOptions.map((option) => (
                                <DefaultButton
                                    key={option.key}
                                    text={option.text}
                                    iconProps={{ iconName: option.iconName }}
                                    primary={actionsUiViewMode === option.key}
                                    onClick={() => {
                                        setActionsUiViewMode(option.key);
                                        setIsViewCalloutOpen(false);
                                    }}
                                />
                            ))}
                        </Stack>
                    </Stack>
                </Callout>
            )}
            <PrimaryButton
                text="Új információ"
                iconProps={{ iconName: 'Add' }}
                onClick={openPanel}
                disabled={!isEditor}
            />


        </Stack>
    );
};