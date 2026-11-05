import * as React from 'react';
import {
    Dialog,
    DialogFooter,
    DialogType,
    Panel, PanelType, PrimaryButton, DefaultButton,
    Stack, TextField, Dropdown, IDropdownOption,
    DatePicker, DayOfWeek, IDatePickerStrings,
    Label, Link, MessageBar, MessageBarType, Text,
} from '@fluentui/react';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { useGlobalState } from '../store/GlobalStore';
import { IAkcio, defaultAkcio } from '../models/IAkcio';
import pnpListsServices from '../services/pnpListServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './Akciok.module.scss';
import { TiptapEditor } from './TiptapEditor';
import { ShopsSelectionType, TaskShopsSection } from './TaskShopsSection';
import { RelatedFilesSection } from './RelatedFilesSection';

export interface IAkcioPanelProps {
    context: WebPartContext;
}

const datePickerStrings: IDatePickerStrings = {
    months: ['Január', 'Február', 'Március', 'Április', 'Május', 'Június', 'Július', 'Augusztus', 'Szeptember', 'Október', 'November', 'December'],
    shortMonths: ['Jan', 'Feb', 'Már', 'Ápr', 'Máj', 'Jún', 'Júl', 'Aug', 'Sze', 'Okt', 'Nov', 'Dec'],
    days: ['Vasárnap', 'Hétfő', 'Kedd', 'Szerda', 'Csütörtök', 'Péntek', 'Szombat'],
    shortDays: ['V', 'H', 'K', 'Sz', 'Cs', 'P', 'Szo'],
    goToToday: 'Ma',
    prevMonthAriaLabel: 'Előző hónap',
    nextMonthAriaLabel: 'Következő hónap',
    prevYearAriaLabel: 'Előző év',
    nextYearAriaLabel: 'Következő év',
    closeButtonAriaLabel: 'Bezárás',
    monthPickerHeaderAriaLabel: '{0}, a hónap kiválasztásához kattintson ide',
    yearPickerHeaderAriaLabel: '{0}, az év kiválasztásához kattintson ide',
    isRequiredErrorMessage: 'A mező kitöltése kötelező.',
    invalidInputErrorMessage: 'Érvénytelen dátumformátum.',
};

const statusOptions: IDropdownOption[] = [
    { key: 'Piszkozat', text: 'Piszkozat' },
    { key: 'Közzétéve', text: 'Közzétéve' },
];

const relatedShopsTypeOptions: IDropdownOption[] = [
    { key: 'all', text: 'Összes bolt' },
    { key: 'group', text: 'Csoport' },
    { key: 'selected', text: 'Kiválasztott' },
];

const formatDate = (date: Date | undefined): string =>
    date ? date.toLocaleDateString('hu-HU') : '—';

const formatSavedAt = (date: Date): string => {
    const pad = (value: number): string => String(value).padStart(2, '0');

    return `${date.getFullYear()}.${pad(date.getMonth() + 1)}.${pad(date.getDate())}. ${pad(date.getHours())}:${pad(date.getMinutes())}`;
};

const toDisplayText = (value: string | undefined): string =>
    value && value.trim().length > 0 ? value : '—';

const getOptionText = (options: IDropdownOption[], key: string): string =>
    options.find(o => String(o.key) === key)?.text ?? key;

const parseRelatedShops = (value: string): string[] =>
    value
        .split(/[\n,;]+/)
        .map(item => item.trim())
        .filter(item => item.length > 0);

const parseRelatedProducts = (value: string): string[] =>
    value
        .split(/[\n,;]+/)
        .map(item => item.trim())
        .filter(item => item.length > 0);

const mergeRelatedProducts = (existingValue: string, incomingValues: string[]): string => {
    const normalizedIncoming = incomingValues
        .map((item) => item.trim())
        .filter((item) => item.length > 0);

    if (!normalizedIncoming.length) {
        return existingValue;
    }

    const merged = parseRelatedProducts(existingValue);
    normalizedIncoming.forEach((item) => {
        if (merged.indexOf(item) === -1) {
            merged.push(item);
        }
    });

    return merged.join(', ');
};

const serializeRelatedShops = (values: string[]): string => values.join(',');

const toDefaultSelectedUser = (email?: string, title?: string): string => {
    if (email && email.trim()) {
        return email;
    }

    if (title && title.trim()) {
        return `/${title.trim()}`;
    }

    return '';
};

const SCRIPT_PROTOCOL = ['java', 'script:'].join('');
const MAX_INLINE_IMAGE_BASE64_LENGTH = 300000;
const MAX_TOTAL_INLINE_IMAGE_BASE64_LENGTH = 900000;

const extractInlineImageBase64Sizes = (html: string): number[] => {
    const sizes: number[] = [];
    const regex = /<img[^>]+src=["']data:image\/[^;]+;base64,([^"']+)["'][^>]*>/gi;
    let match = regex.exec(html);

    while (match) {
        sizes.push(match[1].length);
        match = regex.exec(html);
    }

    return sizes;
};

const getInlineImageWarning = (descriptionHtml: string): string => {
    const allSizes = extractInlineImageBase64Sizes(descriptionHtml);

    if (!allSizes.length) {
        return '';
    }

    const maxImageSize = Math.max(...allSizes);
    const totalSize = allSizes.reduce((sum, size) => sum + size, 0);

    if (maxImageSize > MAX_INLINE_IMAGE_BASE64_LENGTH) {
        return 'A beillesztett kép túl nagy base64 tartalommal került be a leírásba. Kérlek használj kisebb képet vagy külső kép URL-t, különben SharePoint mentésnél elveszhet a kép.';
    }

    if (totalSize > MAX_TOTAL_INLINE_IMAGE_BASE64_LENGTH) {
        return 'A beágyazott képek összmérete túl nagy. Kérlek csökkentsd a képek méretét vagy használj hivatkozott képeket, különben SharePoint mentésnél elveszhetnek.';
    }

    return '';
};

const sanitizeHtml = (html: string | undefined): string => {
    if (!html || !html.trim()) {
        return '';
    }

    const template = document.createElement('template');
    template.innerHTML = html;

    template.content.querySelectorAll('script, style, iframe, object, embed').forEach(node => node.remove());

    template.content.querySelectorAll('*').forEach((element) => {
        Array.from(element.attributes).forEach((attr) => {
            const name = attr.name.toLowerCase();
            const value = attr.value.trim().toLowerCase();

            if (name.startsWith('on')) {
                element.removeAttribute(attr.name);
            }

            if ((name === 'href' || name === 'src') && value.startsWith(SCRIPT_PROTOCOL)) {
                element.removeAttribute(attr.name);
            }
        });
    });

    template.content.querySelectorAll('a[href]').forEach((link) => {
        link.setAttribute('target', '_blank');
        link.setAttribute('rel', 'noopener noreferrer');
        // Stops the SharePoint page router from handling the click in the current window.
        link.setAttribute('data-interception', 'off');
    });

    return template.innerHTML;
};

const openClickedImageInNewWindow = (event: React.MouseEvent<HTMLElement>): void => {
    const clickedElement = event.target;

    if (!(clickedElement instanceof HTMLImageElement) || clickedElement.closest('a')) {
        return;
    }

    const source = clickedElement.currentSrc || clickedElement.src;

    if (!source.startsWith('data:')) {
        window.open(source, '_blank', 'noopener');
        return;
    }

    // Browsers block top-level navigation to data: URLs, so the image is shown in an empty window instead.
    // An <img> element also keeps any script inside an embedded SVG from running.
    const imageWindow = window.open('', '_blank');

    if (!imageWindow) {
        return;
    }

    const image = imageWindow.document.createElement('img');
    image.src = source;
    image.alt = clickedElement.alt;
    image.style.maxWidth = '100%';
    imageWindow.document.body.appendChild(image);
};

const buildTasksPageUrl = (tasksListPage: string, actionId: number): string => {
    const trimmedValue = tasksListPage.trim();
    if (!trimmedValue) {
        return '';
    }

    try {
        const baseUrl = trimmedValue.startsWith('http://') || trimmedValue.startsWith('https://')
            ? trimmedValue
            : new URL(trimmedValue, window.location.origin).toString();

        const url = new URL(baseUrl);
        url.searchParams.set('actionId', String(actionId));
        return url.toString();
    } catch {
        const separator = trimmedValue.indexOf('?') === -1 ? '?' : '&';
        return `${trimmedValue}${separator}actionId=${encodeURIComponent(String(actionId))}`;
    }
};

export const AkcioPanel: React.FunctionComponent<IAkcioPanelProps> = ({ context }) => {
    const {
        isPanelOpen,
        closePanel,
        actionsListId,
        tasksListPage,
        shopListItems,
        triggerRefresh,
        selectedActionId,
        panelMode,
        setPanelMode,
        openActionPanel,
        isEditor,
    } = useGlobalState();
    const [form, setForm] = React.useState<IAkcio>({ ...defaultAkcio });
    const [responsiblePersonTitle, setResponsiblePersonTitle] = React.useState<string>('');
    const [responsiblePersonEmail, setResponsiblePersonEmail] = React.useState<string>('');
    const [responsiblePersonDefaultUser, setResponsiblePersonDefaultUser] = React.useState<string>('');
    const [isLoadingItem, setIsLoadingItem] = React.useState(false);
    const [isSaving, setIsSaving] = React.useState(false);
    const [isPublishDialogOpen, setIsPublishDialogOpen] = React.useState(false);
    const [isPublishInProgress, setIsPublishInProgress] = React.useState(false);
    const [errorMessage, setErrorMessage] = React.useState<string>('');
    const [warningMessage, setWarningMessage] = React.useState<string>('');
    const [actionTypeOptions, setActionTypeOptions] = React.useState<IDropdownOption[]>([]);
    const [lastSavedAt, setLastSavedAt] = React.useState<Date | undefined>(undefined);
    const justCreatedActionIdRef = React.useRef<number | undefined>(undefined);

    const effectiveActionTypeOptions = React.useMemo<IDropdownOption[]>(() => {
        if (!form.actionType || actionTypeOptions.some((option) => String(option.key) === form.actionType)) {
            return actionTypeOptions;
        }

        return [{ key: form.actionType, text: form.actionType }, ...actionTypeOptions];
    }, [actionTypeOptions, form.actionType]);

    const selectedShops = React.useMemo(() => parseRelatedShops(form.relatedShops), [form.relatedShops]);

    const relatedShopsDisplayValue = React.useMemo(() => {
        if (form.relatedShopsType === 'all') {
            return '';
        }

        if (form.relatedShopsType !== 'selected') {
            return form.relatedShops;
        }

        if (!selectedShops.length) {
            return '';
        }

        const shopNameById = new Map(shopListItems.map((shop) => [String(shop.id), shop.title]));
        return selectedShops
            .map((shopId) => shopNameById.get(shopId) || shopId)
            .join(', ');
    }, [form.relatedShopsType, form.relatedShops, selectedShops, shopListItems]);

    const isReadOnly = panelMode === 'view';
    const isCreateMode = panelMode === 'create';

    const closePanelState = (force: boolean = false): void => {
        if (!force && (isSaving || isPublishInProgress || isPublishDialogOpen)) {
            return;
        }

        setForm({ ...defaultAkcio });
        setResponsiblePersonTitle('');
        setResponsiblePersonEmail('');
        setResponsiblePersonDefaultUser('');
        setErrorMessage('');
        setWarningMessage('');
        setLastSavedAt(undefined);
        closePanel();
    };

    const handleDismiss = (): void => {
        closePanelState(false);
    };

    const handleSave = async (formOverride?: IAkcio, allowReadOnlySave?: boolean): Promise<boolean> => {
        if (isSaving) {
            return false;
        }

        const formToSave = formOverride ?? form;

        if (isReadOnly && !allowReadOnlySave) {
            return false;
        }

        if (!formToSave.title.trim()) {
            setErrorMessage('Az információ címe kötelező.');
            return false;
        }

        if (!formToSave.actionType.trim()) {
            setErrorMessage('Az információ típusa kötelező.');
            return false;
        }

        const imageWarning = getInlineImageWarning(formToSave.description);
        if (imageWarning) {
            setWarningMessage(imageWarning);
        } else {
            setWarningMessage('');
        }

        setErrorMessage('');
        setIsSaving(true);
        try {
            if (isCreateMode) {
                const createdActionId = await pnpListsServices.createAction(actionsListId, formToSave);
                justCreatedActionIdRef.current = createdActionId;
                openActionPanel(createdActionId, 'edit');
                await pnpListsServices.setActionStatus(actionsListId, createdActionId, formToSave.status);
            } else if (selectedActionId) {
                await pnpListsServices.updateAction(actionsListId, selectedActionId, formToSave);
            }
            setLastSavedAt(new Date());
            triggerRefresh();
            return true;
        } catch (e) {
            setErrorMessage(`Mentési hiba: ${e instanceof Error ? e.message : String(e)}`);
            return false;
        } finally {
            setIsSaving(false);
        }
    };

    React.useEffect(() => {
        if (!isPanelOpen || !actionsListId) {
            return;
        }

        let isCancelled = false;

        pnpListsServices.getActionTypeChoices(actionsListId)
            .then((choices) => {
                if (isCancelled) {
                    return;
                }

                setActionTypeOptions(choices.map((choice) => ({ key: choice, text: choice })));
            })
            .catch(() => {
                if (!isCancelled) {
                    setActionTypeOptions([]);
                }
            });

        return () => {
            isCancelled = true;
        };
    }, [isPanelOpen, actionsListId]);

    React.useEffect(() => {
        if (!isPanelOpen || panelMode !== 'create') {
            return;
        }

        if (form.actionType || !actionTypeOptions.length) {
            return;
        }

        setForm((currentForm) => ({
            ...currentForm,
            actionType: String(actionTypeOptions[0].key),
        }));
    }, [isPanelOpen, panelMode, form.actionType, actionTypeOptions]);

    React.useEffect(() => {
        if (!isPanelOpen) {
            return;
        }

        if (panelMode === 'create') {
            setForm({ ...defaultAkcio });
            setResponsiblePersonTitle('');
            setResponsiblePersonEmail('');
            setResponsiblePersonDefaultUser('');
            setErrorMessage('');
            setWarningMessage('');
            return;
        }

        if (!selectedActionId) {
            setErrorMessage('A kiválasztott információ nem található.');
            return;
        }

        // The form already holds the item that was just created, and reloading it could overwrite newer typing.
        if (justCreatedActionIdRef.current === selectedActionId) {
            justCreatedActionIdRef.current = undefined;
            return;
        }

        setIsLoadingItem(true);
        setErrorMessage('');
        setWarningMessage('');

        pnpListsServices.getActionById(actionsListId, selectedActionId)
            .then((item) => {
                setForm({
                    title: item.title,
                    actionType: item.actionType,
                    status: item.status,
                    startDate: item.startDate,
                    endDate: item.endDate,
                    publishStartDate: item.publishStartDate,
                    description: item.description,
                    responsiblePersonLoginName: item.responsiblePersonLoginName,
                    responsibleEmail: item.responsibleEmail ?? '',
                    relatedDocuments: item.relatedDocuments,
                    relatedShopsType: item.relatedShopsType,
                    relatedShops: item.relatedShops,
                    relatedProducts: item.relatedProducts,
                    images: item.images,
                });
                setResponsiblePersonTitle(item.responsiblePersonTitle ?? '');
                setResponsiblePersonEmail(item.responsiblePersonEmail ?? '');
                setResponsiblePersonDefaultUser(toDefaultSelectedUser(item.responsiblePersonEmail, item.responsiblePersonTitle));
            })
            .catch((e) => {
                setErrorMessage(`Betöltési hiba: ${e instanceof Error ? e.message : String(e)}`);
            })
            .finally(() => {
                setIsLoadingItem(false);
            });
    }, [isPanelOpen, panelMode, selectedActionId, actionsListId]);

    const peoplePickerContext: IPeoplePickerContext = {
        absoluteUrl: context.pageContext.web.absoluteUrl,
        msGraphClientFactory: context.msGraphClientFactory as any,
        spHttpClient: context.spHttpClient as any,
    };

    const getHeaderText = (): string => {
        if (panelMode === 'create') {
            return 'Új információ';
        }

        if (panelMode === 'edit') {
            return 'Információ szerkesztése';
        }

        return 'Információ megtekintése';
    };

    const getPublishDialogText = (): string => {
        if (form.publishStartDate) {
            return `Az információ ${formatDate(form.publishStartDate)} dátumtól lesz mindenki számára elérhető. Folytatod a közzétételt?`;
        }

        return 'Nincs megadva publikálási dátum, ezért az információ azonnal mindenki számára elérhető lesz. Folytatod a közzétételt?';
    };

    const handlePublishConfirm = async (): Promise<void> => {
        if (isPublishInProgress) {
            return;
        }

        setIsPublishInProgress(true);
        setIsPublishDialogOpen(false);
        const publishedForm: IAkcio = { ...form, status: 'Közzétéve' };
        setForm(publishedForm);
        let isSaved = false;

        try {
            isSaved = await handleSave(publishedForm, true);
        } finally {
            setIsPublishInProgress(false);
        }

        if (isSaved) {
            // Forced: this closure still sees the publish dialog as open, which would block a normal dismiss.
            closePanelState(true);
        }
    };

    const onRenderFooterContent = (): JSX.Element => (
        <Stack tokens={{ childrenGap: 8 }}>
            {warningMessage && (
                <Text variant="small" styles={{ root: { color: '#a4262c' } }}>
                    {warningMessage}
                </Text>
            )}
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                {isEditor && panelMode === 'edit' && (
                    <DefaultButton text="Megtekintő nézet" onClick={() => setPanelMode('view')} disabled={isSaving || isLoadingItem} />
                )}
                {isEditor && panelMode === 'view' && (
                    <DefaultButton text="Szerkesztő nézet" onClick={() => setPanelMode('edit')} disabled={isSaving || isLoadingItem} />
                )}
                {isEditor && !isReadOnly && (
                    <DefaultButton
                        text="Közzététel"
                        iconProps={{ iconName: 'PublishContent' }}
                        onClick={() => setIsPublishDialogOpen(true)}
                        disabled={isSaving || isLoadingItem || isPublishInProgress || form.status === 'Közzétéve'}
                    />
                )}
                {!isReadOnly && (
                    <PrimaryButton text="Mentés" onClick={() => { handleSave().catch(() => undefined); }} disabled={isSaving || isLoadingItem || isPublishInProgress} />
                )}
                <DefaultButton text={isReadOnly ? 'Bezárás' : 'Mégsem'} onClick={handleDismiss} disabled={isSaving || isPublishInProgress} />
                {!isReadOnly && lastSavedAt && (
                    <Text variant="small">Mentve: {formatSavedAt(lastSavedAt)}</Text>
                )}
            </Stack>
        </Stack>
    );

    const renderViewField = (label: string, value: string): JSX.Element => (
        <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="small" className={styles.viewerLabel}>{label}:</Text>
            <Text variant="medium" className={styles.viewerValue}>{toDisplayText(value)}</Text>
        </Stack>
    );

    const renderViewLinkField = (label: string, text: string, href?: string): JSX.Element => {
        const hasLink = !!href && href.trim().length > 0;

        return (
            <Stack tokens={{ childrenGap: 4 }}>
                <Text variant="small" className={styles.viewerLabel}>{label}:</Text>
                {hasLink ? (
                    <Link href={href} target="_blank">{toDisplayText(text)}</Link>
                ) : (
                    <Text variant="medium" className={styles.viewerValue}>{toDisplayText(text)}</Text>
                )}
            </Stack>
        );
    };

    const renderViewHtmlField = (label: string, value: string): JSX.Element => {
        const sanitizedHtml = sanitizeHtml(value);

        return (
            <Stack tokens={{ childrenGap: 4 }}>
                <Text variant="small" className={styles.viewerLabel}>{label}:</Text>
                {sanitizedHtml ? (
                    <div
                        className={styles.viewerHtmlValue}
                        onClick={openClickedImageInNewWindow}
                        dangerouslySetInnerHTML={{ __html: sanitizedHtml }}
                    />
                ) : (
                    <Text variant="medium" className={styles.viewerValue}>—</Text>
                )}
            </Stack>
        );
    };

    const renderViewSection = (title: string, content: JSX.Element | JSX.Element[]): JSX.Element => (
        <Stack className={styles.viewerSection} tokens={{ childrenGap: 10 }}>
            <Text variant="mediumPlus" className={styles.viewerSectionTitle}>{title}</Text>
            {content}
        </Stack>
    );

    const renderEditSection = (title: string, content: JSX.Element | JSX.Element[]): JSX.Element => (
        <Stack className={styles.formSection} tokens={{ childrenGap: 10 }}>
            <Text variant="mediumPlus" className={styles.formSectionTitle}>{title}</Text>
            <Stack tokens={{ childrenGap: 10 }} className={styles.formSectionFields}>
                {content}
            </Stack>
        </Stack>
    );

    const renderReadOnlyView = (): JSX.Element => (
        <Stack tokens={{ childrenGap: 12 }}>
            <Stack className={styles.viewerHero} tokens={{ childrenGap: 8 }}>
                <Text variant="xLarge" className={styles.viewerTitle}>{toDisplayText(form.title)}</Text>
                <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
                    <span className={`${styles.statusBadge} ${form.status === 'Közzétéve' ? styles.statusPublished : styles.statusDraft}`}>
                        {toDisplayText(form.status)}
                    </span>
                    <Text variant="small" className={styles.viewerChip}>{getOptionText(effectiveActionTypeOptions, form.actionType)}</Text>
                </Stack>
            </Stack>

            {renderViewSection('Dátumok', [
                renderViewField('Információ megjelenési dátuma', formatDate(form.startDate)),
                renderViewField('Információ befejező dátuma', formatDate(form.endDate)),
                //renderViewField('Információ közzétételi dátuma', formatDate(form.publishStartDate)),
            ])}

            {renderViewSection('Leírás', [
                renderViewHtmlField('Információ leírása', form.description),
            ])}

            {renderViewSection('Felelős és boltok', [
                renderViewLinkField('Felelős személy', responsiblePersonTitle, responsiblePersonEmail ? `mailto:${responsiblePersonEmail}` : ''),
                renderViewLinkField('Felelős email', form.responsibleEmail, form.responsibleEmail ? `mailto:${form.responsibleEmail}` : ''),
                renderViewField('Kapcsolódó boltok típusa', getOptionText(relatedShopsTypeOptions, form.relatedShopsType)),
                ...(form.relatedShopsType !== 'all' ? [renderViewField('Kapcsolódó boltok', relatedShopsDisplayValue)] : []),
            ])}

            {renderViewSection('Egyéb', [
                ...(tasksListPage && selectedActionId
                    ? [
                        <DefaultButton
                            key="relatedTasksLink"
                            text="Kapcsolódó feladatok"
                            iconProps={{ iconName: 'OpenInNewWindow' }}
                            href={buildTasksPageUrl(tasksListPage, selectedActionId)}
                            target="_blank"
                        />,
                    ]
                    : []),
                <RelatedFilesSection key="relatedFilesReadOnly" actionId={selectedActionId} readOnly={true} />,
            ])}
        </Stack>
    );

    return (
        <>
            <Panel
                isOpen={isPanelOpen}
                onDismiss={handleDismiss}
                isLightDismiss={false}
                type={PanelType.custom}
                customWidth="900px"
                headerText={getHeaderText()}
                onRenderFooterContent={onRenderFooterContent}
                isFooterAtBottom
            >
                <Stack tokens={{ childrenGap: 12 }}>
                {errorMessage && (
                    <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMessage('')}>
                        {errorMessage}
                    </MessageBar>
                )}

                {warningMessage && (
                    <MessageBar messageBarType={MessageBarType.warning} onDismiss={() => setWarningMessage('')}>
                        {warningMessage}
                    </MessageBar>
                )}

                {isLoadingItem && (
                    <MessageBar messageBarType={MessageBarType.info}>
                        Információ adatainak betöltése...
                    </MessageBar>
                )}

                {isReadOnly ? renderReadOnlyView() : (
                    <>
                        {renderEditSection('Alapadatok', [
                            <TextField
                                key="title"
                                label="Információ címe"
                                required
                                value={form.title}
                                onChange={(_, v) => setForm(f => ({ ...f, title: v ?? '' }))}
                            />,
                            <Dropdown
                                key="actionType"
                                label="Információ típusa"
                                options={effectiveActionTypeOptions}
                                selectedKey={form.actionType}
                                disabled={!effectiveActionTypeOptions.length}
                                onChange={(_, o) => setForm(f => ({ ...f, actionType: String(o?.key ?? '') }))}
                            />,
                            <Dropdown
                                key="status"
                                label="Státusz"
                                options={statusOptions}
                                selectedKey={form.status}
                                onChange={(_, o) => setForm(f => ({ ...f, status: String(o?.key ?? '') }))}
                            />,
                        ])}

                        {renderEditSection('Dátumok', [
                            <DatePicker
                                key="startDate"
                                label="Információ megjelenési dátuma"
                                firstDayOfWeek={DayOfWeek.Monday}
                                strings={datePickerStrings}
                                value={form.startDate}
                                onSelectDate={(d) => setForm(f => ({ ...f, startDate: d ?? undefined }))}
                                formatDate={(d) => d ? d.toLocaleDateString('hu-HU') : ''}
                            />,
                            <DatePicker
                                key="endDate"
                                label="Információ befejező dátuma"
                                firstDayOfWeek={DayOfWeek.Monday}
                                strings={datePickerStrings}
                                value={form.endDate}
                                onSelectDate={(d) => setForm(f => ({ ...f, endDate: d ?? undefined }))}
                                formatDate={(d) => d ? d.toLocaleDateString('hu-HU') : ''}
                            />,
                            <DatePicker
                                key="publishStartDate"
                                label="Információ közzétételi dátuma"
                                firstDayOfWeek={DayOfWeek.Monday}
                                strings={datePickerStrings}
                                value={form.publishStartDate}
                                onSelectDate={(d) => setForm(f => ({ ...f, publishStartDate: d ?? undefined }))}
                                formatDate={(d) => d ? d.toLocaleDateString('hu-HU') : ''}
                            />,
                        ])}

                        {renderEditSection('Leírás', [
                            <TiptapEditor
                                key="description"
                                label="Információ leírása"
                                value={form.description}
                                onChange={(value) => setForm(f => ({ ...f, description: value }))}
                                showGoodsSearch={true}
                                onEansDetected={(codes) => {
                                    setForm((currentForm) => ({
                                        ...currentForm,
                                        relatedProducts: mergeRelatedProducts(currentForm.relatedProducts, codes),
                                    }));
                                }}
                            />,
                        ])}

                        {renderEditSection('Felelős és boltok', [
                            <Stack key="responsiblePerson" tokens={{ childrenGap: 6 }}>
                                <Label>Felelős személy</Label>
                                <PeoplePicker
                                    context={peoplePickerContext}
                                    personSelectionLimit={1}
                                    principalTypes={[PrincipalType.User]}
                                    resolveDelay={500}
                                    groupId="570b9851-c9cd-47d9-a204-cbfe99a373ed" // Leave empty to allow selection of any user
                                    searchTextLimit={2}
                                    defaultSelectedUsers={responsiblePersonDefaultUser ? [responsiblePersonDefaultUser] : []}
                                    onChange={(items) => {
                                        const first = items.length > 0 ? (items[0] as any) : undefined;
                                        const loginName = first ? String(first.loginName || first.id || first.secondaryText || '') : '';
                                        const defaultUser = first ? toDefaultSelectedUser(String(first.secondaryText || ''), String(first.text || '')) : '';
                                        setResponsiblePersonDefaultUser(defaultUser);
                                        setResponsiblePersonEmail(first ? String(first.secondaryText || '') : '');
                                        setForm(f => ({ ...f, responsiblePersonLoginName: loginName }));
                                    }}
                                />
                            </Stack>,
                            <TextField
                                key="responsibleEmail"
                                label="Felelős email"
                                value={form.responsibleEmail}
                                onChange={(_, v) => setForm(f => ({ ...f, responsibleEmail: v ?? '' }))}
                            />,
                            <TaskShopsSection
                                key="relatedShopsSection"
                                shopsSelectionType={form.relatedShopsType as ShopsSelectionType}
                                selectedShops={selectedShops}
                                onFormChange={(updates) => {
                                    setForm((currentForm) => ({
                                        ...currentForm,
                                        relatedShopsType: updates.shopsSelectionType ?? currentForm.relatedShopsType,
                                        relatedShops: updates.selectedShops ? serializeRelatedShops(updates.selectedShops) : currentForm.relatedShops,
                                    }));
                                }}
                            />,
                        ])}

                        {renderEditSection('Egyéb', [
                            <TextField
                                key="relatedProducts"
                                label="Kapcsolódó termékek"
                                multiline
                                rows={3}
                                value={form.relatedProducts}
                                onChange={(_, v) => setForm(f => ({ ...f, relatedProducts: v ?? '' }))}
                            />,
                            <RelatedFilesSection key="relatedFilesEditor" actionId={selectedActionId} readOnly={false} />,
                        ])}
                    </>
                )}
                </Stack>
            </Panel>

            <Dialog
                hidden={!isPublishDialogOpen}
                onDismiss={() => {
                    if (!isPublishInProgress) {
                        setIsPublishDialogOpen(false);
                    }
                }}
                modalProps={{
                    isBlocking: true,
                }}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Közzététel megerősítése',
                    subText: getPublishDialogText(),
                }}
            >
                <DialogFooter>
                    <PrimaryButton
                        text={isPublishInProgress ? 'Közzététel folyamatban...' : 'OK'}
                        onClick={() => { handlePublishConfirm().catch(() => undefined); }}
                        disabled={isSaving || isLoadingItem || isPublishInProgress}
                    />
                    <DefaultButton text="Mégsem" onClick={() => setIsPublishDialogOpen(false)} disabled={isSaving || isPublishInProgress} />
                </DialogFooter>
            </Dialog>
        </>
    );
};
