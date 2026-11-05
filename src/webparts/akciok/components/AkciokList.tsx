import * as React from 'react';
import {
    DefaultButton,
    Spinner, SpinnerSize, MessageBar, MessageBarType,
    Stack, Text, Icon,
} from '@fluentui/react';
import { useGlobalState } from '../store/GlobalStore';
import { IAkcioListItem } from '../models/IAkcio';
import pnpListsServices from '../services/pnpListServices';
import styles from './Akciok.module.scss';

const PAGE_SIZE = 20;

const formatDate = (iso: string | null): string =>
    iso ? new Date(iso).toLocaleDateString('hu-HU') : '—';

const getActionTypeColor = (actionTypeColors: Record<string, string>, actionType: string): string => (
    actionTypeColors[actionType] || '#a19f9d'
);

const AkcioCard: React.FunctionComponent<{ item: IAkcioListItem; onOpen: (itemId: number) => void; actionTypeColor: string }> = ({ item, onOpen, actionTypeColor }) => (
    <div
        className={styles.akcioCard}
        style={{ borderLeft: `4px solid ${actionTypeColor}` }}
        role="button"
        tabIndex={0}
        onClick={() => onOpen(item.id)}
        onKeyDown={(ev) => {
            if (ev.key === 'Enter' || ev.key === ' ') {
                ev.preventDefault();
                onOpen(item.id);
            }
        }}
    >
        <div className={styles.akcioCardHeader}>
            <span className={`${styles.statusBadge} ${item.status === 'Közzétéve' ? styles.statusPublished : styles.statusDraft}`}>
                {item.status}
            </span>
            <span className={styles.actionTypeColorDot} style={{ backgroundColor: actionTypeColor }} aria-hidden="true" />
            <span className={styles.actionType}>{item.actionType}</span>
        </div>
        <Text variant="mediumPlus" className={styles.akcioCardTitle}>{item.title}</Text>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }} className={styles.akcioCardDates}>
            <Icon iconName="Calendar" />
            <Text variant="small">Lejárat: {formatDate(item.endDate)}</Text>
        </Stack>
        {item.responsiblePersonTitle && (
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }} className={styles.akcioCardPerson}>
                <Icon iconName="Contact" />
                <Text variant="small">{item.responsiblePersonTitle}</Text>
            </Stack>
        )}
    </div>
);

export interface IAkciokListProps {
    actionTypeColors: Record<string, string>;
}

export const AkciokList: React.FunctionComponent<IAkciokListProps> = ({ actionTypeColors }) => {
    const {
        actionsListId,
        refreshKey,
        isEditor,
        openActionPanel,
        actionSortBy,
        hideExpiredActions,
        hideArchivedActions,
        actionsUiViewMode,
        selectedActionTypes,
        actionSearchText,
        setAvailableActionTypes,
        setSelectedActionTypes,
    } = useGlobalState();
    const [items, setItems] = React.useState<IAkcioListItem[]>([]);
    const [currentPage, setCurrentPage] = React.useState(1);
    const [isLoading, setIsLoading] = React.useState(false);
    const [error, setError] = React.useState<string>('');

    React.useEffect(() => {
        if (!actionsListId) return;
        setCurrentPage(1);
        setIsLoading(true);
        setError('');
        pnpListsServices.getActions(actionsListId)
            .then(setItems)
            .catch(e => setError(e instanceof Error ? e.message : String(e)))
            .finally(() => setIsLoading(false));
    }, [actionsListId, refreshKey]);

    const availableActionTypes = React.useMemo(() => {
        const types = Array.from(new Set(items.map((item) => item.actionType).filter((value) => !!value)));
        types.sort((a, b) => a.localeCompare(b, 'hu-HU'));
        return types;
    }, [items]);

    React.useEffect(() => {
        setAvailableActionTypes(availableActionTypes);

        const nextSelectedActionTypes = selectedActionTypes.filter((type) => availableActionTypes.indexOf(type) !== -1);
        if (nextSelectedActionTypes.length !== selectedActionTypes.length) {
            setSelectedActionTypes(nextSelectedActionTypes);
        }
    }, [availableActionTypes, selectedActionTypes, setAvailableActionTypes, setSelectedActionTypes]);

    React.useEffect(() => {
        setCurrentPage(1);
    }, [actionSortBy, hideExpiredActions, hideArchivedActions, selectedActionTypes, actionSearchText]);

    const filteredItems = React.useMemo(() => {
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const normalizedSearchText = actionSearchText.trim().toLowerCase();
        const isSearchActive = normalizedSearchText.length > 0;

        return items.filter((item) => {
            if (!isEditor && item.status !== 'Közzétéve') {
                return false;
            }

            if (hideArchivedActions && item.status === 'Archive') {
                return false;
            }

            if (isSearchActive) {
                const normalizedTitle = item.title.toLowerCase();
                const normalizedRelatedProducts = (item.relatedProducts || '').toLowerCase();

                return normalizedTitle.indexOf(normalizedSearchText) !== -1
                    || normalizedRelatedProducts.indexOf(normalizedSearchText) !== -1;
            }

            if (selectedActionTypes.length && selectedActionTypes.indexOf(item.actionType) === -1) {
                return false;
            }

            if (hideExpiredActions && item.endDate) {
                const endDate = new Date(item.endDate);
                endDate.setHours(0, 0, 0, 0);
                if (endDate.getTime() < today.getTime()) {
                    return false;
                }
            }

            return true;
        });
    }, [items, isEditor, hideExpiredActions, hideArchivedActions, selectedActionTypes, actionSearchText]);

    const sortedItems = React.useMemo(() => {
        const result = [...filteredItems];

        if (actionSortBy === 'actionTypeAsc') {
            result.sort((a, b) => a.actionType.localeCompare(b.actionType, 'hu-HU'));
            return result;
        }

        if (actionSortBy === 'actionTypeDesc') {
            result.sort((a, b) => b.actionType.localeCompare(a.actionType, 'hu-HU'));
            return result;
        }

        if (actionSortBy === 'responsiblePersonAsc') {
            result.sort((a, b) => a.responsiblePersonTitle.localeCompare(b.responsiblePersonTitle, 'hu-HU'));
            return result;
        }

        if (actionSortBy === 'responsiblePersonDesc') {
            result.sort((a, b) => b.responsiblePersonTitle.localeCompare(a.responsiblePersonTitle, 'hu-HU'));
            return result;
        }

        if (actionSortBy === 'endDateAsc') {
            result.sort((a, b) => {
                const aTime = a.endDate ? new Date(a.endDate).getTime() : 0;
                const bTime = b.endDate ? new Date(b.endDate).getTime() : 0;
                return aTime - bTime;
            });
            return result;
        }

        result.sort((a, b) => {
            const aTime = a.endDate ? new Date(a.endDate).getTime() : 0;
            const bTime = b.endDate ? new Date(b.endDate).getTime() : 0;
            return bTime - aTime;
        });

        return result;
    }, [filteredItems, actionSortBy]);

    const totalPages = Math.ceil(sortedItems.length / PAGE_SIZE);
    const shouldShowPagination = sortedItems.length > PAGE_SIZE;
    const firstItemIndex = (currentPage - 1) * PAGE_SIZE;
    const lastItemIndex = firstItemIndex + PAGE_SIZE;
    const visibleItems = sortedItems.slice(firstItemIndex, lastItemIndex);

    const visibleStart = sortedItems.length ? firstItemIndex + 1 : 0;
    const visibleEnd = Math.min(lastItemIndex, sortedItems.length);

    if (isLoading) {
        return <Stack horizontalAlign="center" style={{ padding: '2em' }}><Spinner size={SpinnerSize.large} label="Információk betöltése..." /></Stack>;
    }

    if (error) {
        return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
    }

    if (!sortedItems.length) {
        return <Stack horizontalAlign="center" style={{ padding: '2em' }}><Text variant="medium">Nincs megjeleníthető információ.</Text></Stack>;
    }

    const handleOpenCard = (itemId: number): void => {
        openActionPanel(itemId, isEditor ? 'edit' : 'view');
    };

    return (
        <Stack tokens={{ childrenGap: 10 }}>
            {actionsUiViewMode === 'cards' ? (
                <div className={styles.akcioCardGrid}>
                    {visibleItems.map(item => (
                        <AkcioCard
                            key={item.id}
                            item={item}
                            onOpen={handleOpenCard}
                            actionTypeColor={getActionTypeColor(actionTypeColors, item.actionType)}
                        />
                    ))}
                </div>
            ) : (
                <div className={styles.akcioListTableWrap}>
                    <table className={styles.akcioListTable}>
                        <thead>
                            <tr>
                                <th scope="col">Információ</th>
                                <th scope="col">Típus</th>
                                <th scope="col">Státusz</th>
                                <th scope="col">Lejárat</th>
                                <th scope="col">Felelős</th>
                            </tr>
                        </thead>
                        <tbody>
                            {visibleItems.map(item => (
                                <tr key={item.id}>
                                    <td>
                                        <button
                                            type="button"
                                            className={styles.akcioListTitle}
                                            onClick={() => handleOpenCard(item.id)}
                                        >
                                            {item.title}
                                        </button>
                                    </td>
                                    <td>{item.actionType}</td>
                                    <td>
                                        <span className={`${styles.statusBadge} ${item.status === 'Közzétéve' ? styles.statusPublished : styles.statusDraft}`}>
                                            {item.status}
                                        </span>
                                    </td>
                                    <td>{formatDate(item.endDate)}</td>
                                    <td>{item.responsiblePersonTitle || '—'}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            )}

            {shouldShowPagination && (
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap tokens={{ childrenGap: 8 }}>
                    <Text variant="small">{visibleStart}-{visibleEnd} / {sortedItems.length}</Text>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <DefaultButton
                            text="Előző"
                            disabled={currentPage <= 1}
                            onClick={() => setCurrentPage((page) => Math.max(1, page - 1))}
                        />
                        <Text variant="small">{currentPage}. / {totalPages}. oldal</Text>
                        <DefaultButton
                            text="Következő"
                            disabled={currentPage >= totalPages}
                            onClick={() => setCurrentPage((page) => Math.min(totalPages, page + 1))}
                        />
                    </Stack>
                </Stack>
            )}
        </Stack>
    );
};
