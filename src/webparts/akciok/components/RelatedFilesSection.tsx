import * as React from 'react';
import { DefaultButton, Dialog, DialogFooter, DialogType, Label, Link, MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize, Stack, Text } from '@fluentui/react';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import styles from './Akciok.module.scss';
import pnpListsServices from '../services/pnpListServices';
import { IRelatedFileItem } from '../models/IRelatedFile';
import { useGlobalState } from '../store/GlobalStore';

export interface IRelatedFilesSectionProps {
    actionId?: number;
    readOnly: boolean;
}

const imageExtensions = ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg', '.bmp', '.tiff', '.avif'];

const fileAccepts = [
    '.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg', '.bmp', '.tiff', '.avif',
    '.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.txt', '.csv', '.zip',
];
const brokenFolderIconUrlPart = 'assets/item-types/20/folder.svg';
const fallbackFolderIconDataUrl = 'data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 20 20%22 fill=%22none%22%3E%3Cpath d=%22M2.5 5.5c0-1.1.9-2 2-2h3.3c.4 0 .8.2 1.1.5l1.1 1.1c.3.3.7.4 1 .4h4.5c1.1 0 2 .9 2 2V8H2.5v-2.5Z%22 fill=%22%23f8d26b%22/%3E%3Cpath d=%22M2.5 8h15v6.5c0 1.1-.9 2-2 2h-11c-1.1 0-2-.9-2-2V8Z%22 fill=%22%23d59f00%22/%3E%3Crect x=%222.5%22 y=%228%22 width=%2215%22 height=%228.5%22 rx=%222%22 stroke=%22%23a26f00%22 stroke-width=%221%22/%3E%3C/svg%3E';

const toLower = (value: string): string => value.toLowerCase();

const isImage = (fileName: string): boolean => imageExtensions.some(ext => toLower(fileName).endsWith(ext));

const toReadableSize = (size: number): string => {
    if (size < 1024) {
        return `${size} B`;
    }

    if (size < 1024 * 1024) {
        return `${(size / 1024).toFixed(1)} KB`;
    }

    return `${(size / (1024 * 1024)).toFixed(1)} MB`;
};

export const RelatedFilesSection: React.FunctionComponent<IRelatedFilesSectionProps> = ({ actionId, readOnly }) => {
    const { context, relatedDocumentLibraryId } = useGlobalState();
    const [isFilePickerOpen, setIsFilePickerOpen] = React.useState(false);
    const [files, setFiles] = React.useState<IRelatedFileItem[]>([]);
    const [isLoading, setIsLoading] = React.useState(false);
    const [isUploading, setIsUploading] = React.useState(false);
    const [deletingFileUrl, setDeletingFileUrl] = React.useState<string | undefined>(undefined);
    const [pendingDeleteFile, setPendingDeleteFile] = React.useState<IRelatedFileItem | undefined>(undefined);
    const [error, setError] = React.useState('');

    const canUseLibrary = !!relatedDocumentLibraryId && !!actionId;

    const loadFiles = React.useCallback(async () => {
        if (!canUseLibrary || !actionId) {
            setFiles([]);
            return;
        }

        setIsLoading(true);
        setError('');

        try {
            const loadedFiles = await pnpListsServices.getActionRelatedFiles(relatedDocumentLibraryId, actionId);
            setFiles(loadedFiles);
        } catch (e) {
            setError(e instanceof Error ? e.message : String(e));
        } finally {
            setIsLoading(false);
        }
    }, [canUseLibrary, relatedDocumentLibraryId, actionId]);

    React.useEffect(() => {
        loadFiles().catch(() => undefined);
    }, [loadFiles]);

    React.useEffect(() => {
        if (!isFilePickerOpen) {
            return;
        }

        const patchFolderIcons = (): void => {
            const images = Array.from(document.querySelectorAll('img')) as HTMLImageElement[];
            images.forEach((image) => {
                if (image.dataset.folderIconPatched === '1') {
                    return;
                }

                if (image.src.indexOf(brokenFolderIconUrlPart) === -1) {
                    return;
                }

                image.src = fallbackFolderIconDataUrl;
                image.dataset.folderIconPatched = '1';
            });
        };

        patchFolderIcons();

        const observer = new MutationObserver(() => {
            patchFolderIcons();
        });

        observer.observe(document.body, { childList: true, subtree: true });

        return () => {
            observer.disconnect();
        };
    }, [isFilePickerOpen]);

    const handlePickerSave = async (selectedFiles: IFilePickerResult[]): Promise<void> => {
        setIsFilePickerOpen(false);

        if (!selectedFiles.length || !actionId) {
            return;
        }

        setIsUploading(true);
        setError('');

        try {
            for (const selectedFile of selectedFiles) {
                const fileBlob = await selectedFile.downloadFileContent();
                await pnpListsServices.uploadActionRelatedFile(
                    relatedDocumentLibraryId,
                    actionId,
                    selectedFile.fileName,
                    fileBlob,
                );
            }

            await loadFiles();
        } catch (e) {
            setError(e instanceof Error ? e.message : String(e));
        } finally {
            setIsUploading(false);
        }
    };

    const handleDeleteFile = async (file: IRelatedFileItem): Promise<void> => {
        if (!actionId || !relatedDocumentLibraryId || readOnly) {
            return;
        }

        setPendingDeleteFile(file);
    };

    const confirmDeleteFile = async (): Promise<void> => {
        if (!actionId || !relatedDocumentLibraryId || readOnly || !pendingDeleteFile) {
            return;
        }

        setDeletingFileUrl(pendingDeleteFile.serverRelativeUrl);
        setError('');

        try {
            await pnpListsServices.deleteActionRelatedFile(relatedDocumentLibraryId, actionId, pendingDeleteFile.serverRelativeUrl);
            await loadFiles();
        } catch (e) {
            setError(e instanceof Error ? e.message : String(e));
        } finally {
            setDeletingFileUrl(undefined);
            setPendingDeleteFile(undefined);
        }
    };

    const imageFiles = files.filter(file => isImage(file.name));
    const documentFiles = files.filter(file => !isImage(file.name));
    const shouldRenderImageBlock = readOnly || imageFiles.length > 0;
    const shouldRenderDocumentBlock = true;

    return (
        <Stack tokens={{ childrenGap: 10 }}>
            {!relatedDocumentLibraryId && (
                <MessageBar messageBarType={MessageBarType.warning}>
                    A kapcsolódó dokumentumtár nincs beállítva a webpart tulajdonságai között.
                </MessageBar>
            )}

            {!actionId && (
                <MessageBar messageBarType={MessageBarType.info}>
                    Új információ esetén előbb mentsd az információt, utána lehet dokumentumot és képet feltölteni.
                </MessageBar>
            )}

            {error && (
                <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError('')}>
                    {error}
                </MessageBar>
            )}

            {(isLoading || isUploading) && (
                <Spinner size={SpinnerSize.small} label={isUploading ? 'Fájlok feltöltése...' : 'Fájlok betöltése...'} />
            )}

            {shouldRenderImageBlock && (
                <Stack className={styles.contentBox} tokens={{ childrenGap: 8 }}>
                    <Label>Képek</Label>
                    {imageFiles.length ? (
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10 }}>
                            {imageFiles.map((file) => (
                                <div
                                    key={file.serverRelativeUrl}
                                    style={{ width: 120, display: 'flex', flexDirection: 'column', gap: 4 }}
                                >
                                    <Link href={file.serverRelativeUrl} target="_blank" data-interception="off" title={file.name}>
                                        <img
                                            src={file.serverRelativeUrl}
                                            alt={file.name}
                                            style={{
                                                display: 'block',
                                                width: 100,
                                                height: 100,
                                                maxWidth: 100,
                                                maxHeight: 100,
                                                objectFit: 'cover',
                                                borderRadius: 8,
                                                borderWidth: 1,
                                                borderStyle: 'solid',
                                                borderColor: '#edebe9',
                                                backgroundColor: '#faf9f8',
                                            }}
                                            loading="lazy"
                                        />
                                    </Link>
                                    <Link
                                        href={file.serverRelativeUrl}
                                        target="_blank"
                                        data-interception="off"
                                        style={{ width: 100, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}
                                    >
                                        {file.name}
                                    </Link>
                                    <Text variant="small" className={styles.relatedFileMeta}>{toReadableSize(file.size)}</Text>
                                    {!readOnly && (
                                        <DefaultButton
                                            text="Törlés"
                                            onClick={() => handleDeleteFile(file).catch(() => undefined)}
                                            disabled={!!deletingFileUrl || isUploading || isLoading}
                                        />
                                    )}
                                </div>
                            ))}
                        </div>
                    ) : <Text className={styles.viewerValue}>Nincs feltöltött kép.</Text>}
                </Stack>
            )}

            {shouldRenderDocumentBlock && (
                <Stack className={styles.contentBox} tokens={{ childrenGap: 8 }}>
                    <Label>Dokumentumok</Label>
                    {documentFiles.length ? documentFiles.map((file) => (
                        <Stack horizontal horizontalAlign="space-between" key={file.serverRelativeUrl} className={styles.relatedFileItem}>
                            <Link href={file.serverRelativeUrl} target="_blank" data-interception="off">{file.name}</Link>
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                                <Text variant="small" className={styles.relatedFileMeta}>{toReadableSize(file.size)}</Text>
                                {!readOnly && (
                                    <DefaultButton
                                        text="Törlés"
                                        onClick={() => handleDeleteFile(file).catch(() => undefined)}
                                        disabled={!!deletingFileUrl || isUploading || isLoading}
                                    />
                                )}
                            </Stack>
                        </Stack>
                    )) : <Text className={styles.viewerValue}>Nincs feltöltött dokumentum.</Text>}

                    {!readOnly && (
                        <DefaultButton
                            text={isUploading ? 'Feltöltés folyamatban...' : 'Fájl feltöltése'}
                            onClick={() => setIsFilePickerOpen(true)}
                            disabled={!canUseLibrary || isUploading || isLoading}
                        />
                    )}
                </Stack>
            )}

            {context && (
                <FilePicker
                    context={context as any}
                    hidden={true}
                    isPanelOpen={isFilePickerOpen}
                    onSave={handlePickerSave}
                    onCancel={() => setIsFilePickerOpen(false)}
                    accepts={fileAccepts}
                    hideWebSearchTab={true}
                    hideStockImages={true}
                    hideOrganisationalAssetTab={true}
                    hideOneDriveTab={true}
                    hideRecentTab={true}
                />
            )}

            <Dialog
                hidden={!pendingDeleteFile}
                onDismiss={() => setPendingDeleteFile(undefined)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Fájl törlése',
                    subText: pendingDeleteFile
                        ? `Biztosan törölni szeretnéd ezt a fájlt?\n\n${pendingDeleteFile.name}`
                        : '',
                }}
            >
                <DialogFooter>
                    <PrimaryButton
                        text="Igen, törlöm"
                        onClick={() => confirmDeleteFile().catch(() => undefined)}
                        disabled={!!deletingFileUrl || isUploading || isLoading}
                    />
                    <DefaultButton
                        text="Mégse"
                        onClick={() => setPendingDeleteFile(undefined)}
                        disabled={!!deletingFileUrl}
                    />
                </DialogFooter>
            </Dialog>
        </Stack>
    );
};
