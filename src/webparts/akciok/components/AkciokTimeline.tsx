import * as React from 'react';
import { DefaultButton, Stack, Text } from '@fluentui/react';
import { IAkcioListItem } from '../models/IAkcio';
import styles from './Akciok.module.scss';

export interface IAkciokTimelineProps {
    items: IAkcioListItem[];
    onOpen: (itemId: number) => void;
    actionTypeColors: Record<string, string>;
}

type TimelineItem = IAkcioListItem & {
    start: Date;
    end: Date;
};

const atStartOfDay = (value: Date): Date => {
    const next = new Date(value);
    next.setHours(0, 0, 0, 0);
    return next;
};

const atEndOfDay = (value: Date): Date => {
    const next = new Date(value);
    next.setHours(23, 59, 59, 999);
    return next;
};

const normalizeItem = (item: IAkcioListItem): TimelineItem | null => {
    if (!item.startDate) {
        return null;
    }

    const start = atStartOfDay(new Date(item.startDate));
    const end = item.endDate ? atEndOfDay(new Date(item.endDate)) : atEndOfDay(new Date(item.startDate));

    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
        return null;
    }

    return {
        ...item,
        start,
        end: end.getTime() >= start.getTime() ? end : start,
    };
};

const formatDate = (value: Date): string => value.toLocaleDateString('hu-HU');

const differenceInDays = (start: Date, end: Date): number => {
    const oneDay = 24 * 60 * 60 * 1000;
    return Math.max(1, Math.floor((atStartOfDay(end).getTime() - atStartOfDay(start).getTime()) / oneDay) + 1);
};

const getActionTypeColor = (actionTypeColors: Record<string, string>, actionType: string): string => actionTypeColors[actionType] || '#a19f9d';

export const AkciokTimeline: React.FunctionComponent<IAkciokTimelineProps> = ({ items, onOpen, actionTypeColors }) => {
    const [windowAnchor, setWindowAnchor] = React.useState<Date>(atStartOfDay(new Date()));

    const timelineItems = React.useMemo(() => {
        const normalized = items
            .map(normalizeItem)
            .filter((item): item is TimelineItem => !!item)
            .sort((a, b) => a.start.getTime() - b.start.getTime());

        return normalized;
    }, [items]);

    const windowStart = React.useMemo(() => atStartOfDay(windowAnchor), [windowAnchor]);
    const windowEnd = React.useMemo(() => {
        const end = new Date(windowStart);
        end.setDate(end.getDate() + 27);
        return atEndOfDay(end);
    }, [windowStart]);

    const visibleItems = React.useMemo(() => (
        timelineItems.filter((item) => item.start.getTime() <= windowEnd.getTime() && item.end.getTime() >= windowStart.getTime())
    ), [timelineItems, windowStart, windowEnd]);

    const totalDays = React.useMemo(() => differenceInDays(windowStart, windowEnd), [windowStart, windowEnd]);

    const getLeft = React.useCallback((item: TimelineItem): number => {
        const clippedStart = Math.max(item.start.getTime(), windowStart.getTime());
        return ((clippedStart - windowStart.getTime()) / (24 * 60 * 60 * 1000)) / totalDays * 100;
    }, [windowStart, totalDays]);

    const getWidth = React.useCallback((item: TimelineItem): number => {
        const clippedStart = Math.max(item.start.getTime(), windowStart.getTime());
        const clippedEnd = Math.min(item.end.getTime(), windowEnd.getTime());
        const daySpan = Math.max(1, Math.floor((clippedEnd - clippedStart) / (24 * 60 * 60 * 1000)) + 1);
        return (daySpan / totalDays) * 100;
    }, [windowEnd, windowStart, totalDays]);

    if (!timelineItems.length) {
        return (
            <Stack horizontalAlign="center" style={{ padding: '2em' }}>
                <Text variant="medium">Nincs idővonalon megjeleníthető információ.</Text>
            </Stack>
        );
    }

    return (
        <Stack tokens={{ childrenGap: 12 }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center" className={styles.timelineToolbar}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    <DefaultButton
                        iconProps={{ iconName: 'ChevronLeft' }}
                        text="Előző 4 hét"
                        onClick={() => {
                            const next = new Date(windowStart);
                            next.setDate(next.getDate() - 28);
                            setWindowAnchor(next);
                        }}
                    />
                    <DefaultButton
                        text="Aktuális"
                        onClick={() => setWindowAnchor(atStartOfDay(new Date()))}
                    />
                    <DefaultButton
                        iconProps={{ iconName: 'ChevronRight' }}
                        text="Következő 4 hét"
                        onClick={() => {
                            const next = new Date(windowStart);
                            next.setDate(next.getDate() + 28);
                            setWindowAnchor(next);
                        }}
                    />
                </Stack>
                <Text variant="mediumPlus" className={styles.timelineRangeLabel}>
                    {formatDate(windowStart)} - {formatDate(windowEnd)}
                </Text>
            </Stack>

            <div className={styles.timelineBoard}>
                {visibleItems.length === 0 && (
                    <Stack horizontalAlign="center" style={{ padding: '1.5em' }}>
                        <Text variant="medium">Nincs információ ebben az időablakban.</Text>
                    </Stack>
                )}

                {visibleItems.map((item) => {
                    const actionTypeColor = getActionTypeColor(actionTypeColors, item.actionType);

                    return (
                        <div key={item.id} className={styles.timelineRow}>
                            <button
                                type="button"
                                className={styles.timelineMetaButton}
                                style={{ borderLeft: `4px solid ${actionTypeColor}` }}
                                onClick={() => onOpen(item.id)}
                                title={item.title}
                            >
                                <Text variant="smallPlus" className={styles.timelineTitle}>{item.title}</Text>
                                <Text variant="small" className={styles.timelineMeta}>
                                    <span className={styles.actionTypeColorDot} style={{ backgroundColor: actionTypeColor }} aria-hidden="true" />
                                    {item.actionType} · {formatDate(item.start)} - {formatDate(item.end)}
                                </Text>
                            </button>

                            <div className={styles.timelineTrack}>
                                <button
                                    type="button"
                                    className={[
                                        styles.timelineBar,
                                        item.status === 'Közzétéve' ? styles.timelineBarPublished : styles.timelineBarDraft,
                                    ].join(' ')}
                                    style={{
                                        left: `${getLeft(item)}%`,
                                        width: `${getWidth(item)}%`,
                                        borderColor: actionTypeColor,
                                    }}
                                    onClick={() => onOpen(item.id)}
                                    title={`${item.title} (${item.status})`}
                                >
                                    <span className={styles.timelineBarLabel}>{item.title}</span>
                                </button>
                            </div>
                        </div>
                    );
                })}
            </div>
        </Stack>
    );
};
