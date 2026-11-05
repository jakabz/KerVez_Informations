import * as React from 'react';
import { DefaultButton, Icon, Stack, Text } from '@fluentui/react';
import { IAkcioListItem } from '../models/IAkcio';
import styles from './Akciok.module.scss';

export interface IAkciokCalendarProps {
    items: IAkcioListItem[];
    onOpen: (itemId: number) => void;
    actionTypeColors: Record<string, string>;
}

const WEEKDAY_LABELS = ['H', 'K', 'Sze', 'Cs', 'P', 'Szo', 'V'];
const MAX_VISIBLE_EVENTS = 3;

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

const addDays = (value: Date, days: number): Date => {
    const next = new Date(value);
    next.setDate(next.getDate() + days);
    return next;
};

const getMonthStart = (value: Date): Date => new Date(value.getFullYear(), value.getMonth(), 1);

const getGridStart = (monthStart: Date): Date => {
    const jsDay = monthStart.getDay();
    const mondayOffset = (jsDay + 6) % 7;
    return addDays(monthStart, -mondayOffset);
};

const getGridDays = (monthStart: Date): Date[] => {
    const gridStart = getGridStart(monthStart);
    return Array.from({ length: 42 }, (_, index) => addDays(gridStart, index));
};

const normalizeItemRange = (item: IAkcioListItem): { start: Date; end: Date } | null => {
    if (!item.startDate) {
        return null;
    }

    const start = atStartOfDay(new Date(item.startDate));
    const end = item.endDate ? atEndOfDay(new Date(item.endDate)) : atEndOfDay(new Date(item.startDate));

    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
        return null;
    }

    return { start, end };
};

const containsDate = (range: { start: Date; end: Date }, day: Date): boolean => {
    const startOfDay = atStartOfDay(day).getTime();
    const endOfDay = atEndOfDay(day).getTime();
    return range.start.getTime() <= endOfDay && range.end.getTime() >= startOfDay;
};

const monthTitle = (value: Date): string => value.toLocaleDateString('hu-HU', { year: 'numeric', month: 'long' });
const dayLabel = (value: Date): string => value.toLocaleDateString('hu-HU', { day: 'numeric' });
const getActionTypeColor = (actionTypeColors: Record<string, string>, actionType: string): string => actionTypeColors[actionType] || '#a19f9d';

export const AkciokCalendar: React.FunctionComponent<IAkciokCalendarProps> = ({ items, onOpen, actionTypeColors }) => {
    const [monthCursor, setMonthCursor] = React.useState<Date>(getMonthStart(new Date()));

    const today = React.useMemo(() => atStartOfDay(new Date()), []);

    const days = React.useMemo(() => getGridDays(monthCursor), [monthCursor]);

    const rangesByItemId = React.useMemo(() => {
        const ranges: Record<number, { start: Date; end: Date }> = {};

        for (const item of items) {
            const range = normalizeItemRange(item);
            if (range) {
                ranges[item.id] = range;
            }
        }

        return ranges;
    }, [items]);

    const itemsByDay = React.useMemo(() => {
        const map: Record<string, IAkcioListItem[]> = {};

        for (const day of days) {
            const key = day.toISOString();
            map[key] = items
                .filter((item) => {
                    const range = rangesByItemId[item.id];
                    return !!range && containsDate(range, day);
                })
                .sort((a, b) => {
                    const aStart = rangesByItemId[a.id]?.start.getTime() ?? 0;
                    const bStart = rangesByItemId[b.id]?.start.getTime() ?? 0;
                    return aStart - bStart;
                });
        }

        return map;
    }, [days, items, rangesByItemId]);

    return (
        <Stack tokens={{ childrenGap: 12 }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center" className={styles.calendarToolbar}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    <DefaultButton
                        iconProps={{ iconName: 'ChevronLeft' }}
                        text="Előző hónap"
                        onClick={() => setMonthCursor((current) => new Date(current.getFullYear(), current.getMonth() - 1, 1))}
                    />
                    <DefaultButton
                        text="Aktuális"
                        onClick={() => setMonthCursor(getMonthStart(new Date()))}
                    />
                    <DefaultButton
                        iconProps={{ iconName: 'ChevronRight' }}
                        text="Következő hónap"
                        onClick={() => setMonthCursor((current) => new Date(current.getFullYear(), current.getMonth() + 1, 1))}
                    />
                </Stack>
                <Text variant="large" className={styles.calendarMonthTitle}>{monthTitle(monthCursor)}</Text>
            </Stack>

            <div className={styles.calendarGrid}>
                {WEEKDAY_LABELS.map((label) => (
                    <div key={label} className={styles.calendarWeekdayCell}>
                        <Text variant="smallPlus">{label}</Text>
                    </div>
                ))}

                {days.map((day) => {
                    const dayKey = day.toISOString();
                    const dayItems = itemsByDay[dayKey] ?? [];
                    const isCurrentMonth = day.getMonth() === monthCursor.getMonth();
                    const isToday = day.getTime() === today.getTime();
                    const visibleItems = dayItems.slice(0, MAX_VISIBLE_EVENTS);
                    const hiddenCount = Math.max(0, dayItems.length - visibleItems.length);

                    return (
                        <div
                            key={dayKey}
                            className={[
                                styles.calendarDayCell,
                                !isCurrentMonth ? styles.calendarOutOfMonth : '',
                                isToday ? styles.calendarToday : '',
                            ].filter((value) => !!value).join(' ')}
                        >
                            <Stack tokens={{ childrenGap: 6 }}>
                                <Text variant="small" className={styles.calendarDayNumber}>{dayLabel(day)}</Text>
                                <Stack tokens={{ childrenGap: 4 }}>
                                    {visibleItems.map((item) => {
                                        const actionTypeColor = getActionTypeColor(actionTypeColors, item.actionType);

                                        return (
                                            <button
                                                key={item.id}
                                                className={[
                                                    styles.calendarEvent,
                                                    item.status === 'Közzétéve' ? styles.calendarEventPublished : styles.calendarEventDraft,
                                                ].join(' ')}
                                                style={{ borderLeft: `3px solid ${actionTypeColor}` }}
                                                onClick={() => onOpen(item.id)}
                                                type="button"
                                                title={`${item.title} (${item.actionType})`}
                                            >
                                                <span className={styles.actionTypeColorDot} style={{ backgroundColor: actionTypeColor }} aria-hidden="true" />
                                                <Icon iconName="Calendar" className={styles.calendarEventIcon} />
                                                <span className={styles.calendarEventText}>{item.title}</span>
                                            </button>
                                        );
                                    })}
                                    {hiddenCount > 0 && (
                                        <Text variant="small" className={styles.calendarMoreEvents}>
                                            +{hiddenCount} további
                                        </Text>
                                    )}
                                </Stack>
                            </Stack>
                        </div>
                    );
                })}
            </div>
        </Stack>
    );
};
