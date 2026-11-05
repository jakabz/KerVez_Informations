import * as React from 'react';
import { ComboBox, DefaultButton, IComboBoxOption, IconButton, Label, MessageBar, MessageBarType, Modal, Stack, Text, TextField } from '@fluentui/react';
import { EditorContent, useEditor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import Image from '@tiptap/extension-image';
import TextAlign from '@tiptap/extension-text-align';
import { Table } from '@tiptap/extension-table';
import TableRow from '@tiptap/extension-table-row';
import TableCell from '@tiptap/extension-table-cell';
import TableHeader from '@tiptap/extension-table-header';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { useGlobalState } from '../store/GlobalStore';
import pnpListsServices from '../services/pnpListServices';
import styles from './TiptapEditor.module.scss';

export interface ITiptapEditorProps {
    label: string;
    value: string;
    onChange: (value: string) => void;
    onEansDetected?: (eans: string[]) => void;
    showGoodsSearch?: boolean;
    disabled?: boolean;
}

const emptyDocument = '<p></p>';
const goodsResultColumnKeys: string[] = ['Cikkszam', 'EANszam', 'Megnevezes', 'Ar', 'AkciosAr'];
const goodsCommentColumnKey = '__Megjegyzes';
const goodsCommentColumnTitle = 'Megjegyzés';
const brokenFolderIconUrlPart = 'assets/item-types/20/folder.svg';
const fallbackFolderIconDataUrl = 'data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 20 20%22 fill=%22none%22%3E%3Cpath d=%22M2.5 5.5c0-1.1.9-2 2-2h3.3c.4 0 .8.2 1.1.5l1.1 1.1c.3.3.7.4 1 .4h4.5c1.1 0 2 .9 2 2V8H2.5v-2.5Z%22 fill=%22%23f8d26b%22/%3E%3Cpath d=%22M2.5 8h15v6.5c0 1.1-.9 2-2 2h-11c-1.1 0-2-.9-2-2V8Z%22 fill=%22%23d59f00%22/%3E%3Crect x=%222.5%22 y=%228%22 width=%2215%22 height=%228.5%22 rx=%222%22 stroke=%22%23a26f00%22 stroke-width=%221%22/%3E%3C/svg%3E';

const escapeHtml = (value: string): string => value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

const toEditorContent = (value: string): string => {
    const trimmedValue = value.trim();

    if (!trimmedValue) {
        return emptyDocument;
    }

    if (trimmedValue.startsWith('<')) {
        return value;
    }

    return `<p>${escapeHtml(value)}</p>`;
};

const normalizeEanCandidate = (value: string): string => value.replace(/\D+/g, '');
const normalizeText = (value: string): string => value
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '');
const normalizeProductCodeCandidate = (value: string): string => value.trim();

const isLikelyEan = (value: string): boolean => /^(\d{8}|\d{12}|\d{13}|\d{14})$/.test(value);

const toHeaderToken = (value: string): string => normalizeText(value);

const extractProductCodesFromTableElement = (table: HTMLTableElement): string[] => {
    const rows = Array.from(table.querySelectorAll('tr'));
    if (!rows.length) {
        return [];
    }

    const firstRowCells = Array.from(rows[0].querySelectorAll('th,td'));
    const headerTokens = firstRowCells.map((cell) => toHeaderToken(cell.textContent || ''));
    const eanColumnIndex = headerTokens.findIndex((header) => header.indexOf('ean') !== -1);
    const cikkszamColumnIndex = headerTokens.findIndex((header) => header.indexOf('cikkszam') !== -1);

    if (eanColumnIndex < 0 && cikkszamColumnIndex < 0) {
        return [];
    }

    const result: string[] = [];

    const tryAddEan = (candidate: string): void => {
        const normalized = normalizeEanCandidate(candidate);
        if (isLikelyEan(normalized) && result.indexOf(normalized) === -1) {
            result.push(normalized);
        }
    };

    const tryAddProductCode = (candidate: string): void => {
        const normalized = normalizeProductCodeCandidate(candidate);
        if (normalized && result.indexOf(normalized) === -1) {
            result.push(normalized);
        }
    };

    const rowsToScan = rows.slice(1);
    rowsToScan.forEach((row) => {
        const cells = Array.from(row.querySelectorAll('th,td'));
        if (!cells.length) {
            return;
        }

        if (eanColumnIndex < cells.length) {
            tryAddEan(cells[eanColumnIndex].textContent || '');
        }

        if (cikkszamColumnIndex >= 0 && cikkszamColumnIndex < cells.length) {
            tryAddProductCode(cells[cikkszamColumnIndex].textContent || '');
        }
    });

    return result;
};

const extractProductCodesFromTableHtml = (html: string): string[] => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const tables = Array.from(doc.querySelectorAll('table'));

    if (!tables.length) {
        return [];
    }

    const collected: string[] = [];
    tables.forEach((table) => {
        const fromTable = extractProductCodesFromTableElement(table as HTMLTableElement);
        fromTable.forEach((code) => {
            if (collected.indexOf(code) === -1) {
                collected.push(code);
            }
        });
    });

    return collected;
};

const parseEansFromInput = (value: string): string[] => value
    .split(/[\n,;\t ]+/)
    .map((item) => normalizeEanCandidate(item.trim()))
    .filter((item, index, items) => item.length > 0 && items.indexOf(item) === index);

const ResizableImage = Image.extend({
    addAttributes() {
        return {
            ...this.parent?.(),
            width: {
                default: null,
                parseHTML: (element) => element.getAttribute('width'),
                renderHTML: (attributes) => {
                    if (!attributes.width) {
                        return {};
                    }

                    return { width: attributes.width };
                },
            },
        };
    },
});

export const TiptapEditor: React.FunctionComponent<ITiptapEditorProps> = ({ label, value, onChange, onEansDetected, showGoodsSearch, disabled }) => {
    const { context, goodsListId } = useGlobalState();
    const [isFilePickerOpen, setIsFilePickerOpen] = React.useState<boolean>(false);
    const [isGoodsSearchOpen, setIsGoodsSearchOpen] = React.useState<boolean>(false);
    const [goodsSearchText, setGoodsSearchText] = React.useState<string>('');
    const [goodsSearchResults, setGoodsSearchResults] = React.useState<any[]>([]);
    const [isSearchingGoods, setIsSearchingGoods] = React.useState<boolean>(false);
    const [goodsSearchError, setGoodsSearchError] = React.useState<string>('');
    const [goodsCopyMessage, setGoodsCopyMessage] = React.useState<string>('');
    const [hasSearchedGoods, setHasSearchedGoods] = React.useState<boolean>(false);
    const [isLoadingGoodsColumns, setIsLoadingGoodsColumns] = React.useState<boolean>(false);
    const [goodsColumns, setGoodsColumns] = React.useState<Array<{ internalName: string; title: string }>>([]);
    const [selectedGoodsColumnKeys, setSelectedGoodsColumnKeys] = React.useState<string[]>([]);
    const lastAutoDetectedEansRef = React.useRef<string>('');
    const latestGoodsSearchRequestIdRef = React.useRef<number>(0);
    const skipNextAutoRefreshRef = React.useRef<boolean>(false);

    const goodsColumnTitleByKey = React.useMemo(() => {
        const dictionary: { [key: string]: string } = {};
        goodsColumns.forEach((column) => {
            dictionary[column.internalName] = column.title;
        });
        return dictionary;
    }, [goodsColumns]);

    const effectiveSelectedGoodsColumnKeys = React.useMemo(() => {
        const availableKeys = goodsColumns.map((column) => column.internalName);
        const validSelectedKeys = selectedGoodsColumnKeys.filter((key) => availableKeys.indexOf(key) !== -1);

        return validSelectedKeys;
    }, [goodsColumns, selectedGoodsColumnKeys]);

    const goodsColumnOptions: IComboBoxOption[] = React.useMemo(() => {
        const selectedKeys = effectiveSelectedGoodsColumnKeys;
        return goodsColumns.map((column) => ({
            key: column.internalName,
            text: column.title,
            selected: selectedKeys.indexOf(column.internalName) !== -1,
        }));
    }, [goodsColumns, effectiveSelectedGoodsColumnKeys]);

    const visibleGoodsColumnKeys: string[] = React.useMemo(() => {
        return effectiveSelectedGoodsColumnKeys.concat(goodsCommentColumnKey);
    }, [effectiveSelectedGoodsColumnKeys]);

    React.useEffect(() => {
        if (!isGoodsSearchOpen) {
            return;
        }

        if (!goodsListId) {
            setGoodsColumns([]);
            setSelectedGoodsColumnKeys([]);
            setGoodsSearchError('A terméklista nincs beállítva a webpart tulajdonságai között.');
            return;
        }

        let isCancelled = false;
        setIsLoadingGoodsColumns(true);

        pnpListsServices.getGoodsSearchColumns(goodsListId)
            .then((columns) => {
                if (isCancelled) {
                    return;
                }

                setGoodsColumns(columns);
                const availableKeys = columns.map((column) => column.internalName);
                setSelectedGoodsColumnKeys((previous) => {
                    const keptKeys = previous.filter((key) => availableKeys.indexOf(key) !== -1);
                    return keptKeys;
                });
            })
            .catch((error) => {
                if (isCancelled) {
                    return;
                }

                console.error('Error loading goods columns:', error);
                setGoodsColumns([]);
                setSelectedGoodsColumnKeys([]);
                setGoodsSearchError('Nem sikerült betölteni a cikklista oszlopait.');
            })
            .finally(() => {
                if (!isCancelled) {
                    setIsLoadingGoodsColumns(false);
                }
            });

        return () => {
            isCancelled = true;
        };
    }, [goodsListId, isGoodsSearchOpen]);

    const editor = useEditor({
        immediatelyRender: false,
        extensions: [
            StarterKit,
            ResizableImage,
            TextAlign.configure({
                types: ['heading', 'paragraph'],
            }),
            Table.configure({
                resizable: true,
            }),
            TableRow,
            TableHeader,
            TableCell,
        ],
        editorProps: {
            handlePaste: (_view, event) => {
                const clipboardEvent = event as ClipboardEvent;
                const html = clipboardEvent.clipboardData ? clipboardEvent.clipboardData.getData('text/html') : '';

                if (!html || html.toLowerCase().indexOf('<table') === -1 || !onEansDetected) {
                    return false;
                }

                const detectedProductCodes = extractProductCodesFromTableHtml(html);
                if (detectedProductCodes.length) {
                    onEansDetected(detectedProductCodes);
                }

                return false;
            },
            attributes: {
                class: styles.editorContent,
            },
        },
        content: toEditorContent(value),
        onUpdate: ({ editor: tiptapEditor }) => {
            const html = tiptapEditor.getHTML();
            const hasText = tiptapEditor.getText().trim().length > 0;
            const hasStructuredContent = /<(img|table|hr|ul|ol|li|blockquote|pre|code|h[1-6])\b/i.test(html);

            onChange(hasText || hasStructuredContent ? html : '');

            if (!onEansDetected) {
                return;
            }

            const currentHtml = html;
            if (currentHtml.toLowerCase().indexOf('<table') === -1) {
                return;
            }

            const autoDetectedProductCodes = extractProductCodesFromTableHtml(currentHtml);
            if (!autoDetectedProductCodes.length) {
                return;
            }

            const signature = autoDetectedProductCodes.join('|');
            if (signature === lastAutoDetectedEansRef.current) {
                return;
            }

            lastAutoDetectedEansRef.current = signature;
            onEansDetected(autoDetectedProductCodes);
        },
    });

    React.useEffect(() => {
        if (!editor) {
            return;
        }

        const nextContent = toEditorContent(value);

        if (editor.getHTML() !== nextContent) {
            editor.commands.setContent(nextContent, { emitUpdate: false });
        }
    }, [editor, value]);

    React.useEffect(() => {
        if (!editor) {
            return;
        }

        editor.setEditable(!disabled);
    }, [editor, disabled]);

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

    const runCommand = (command: () => boolean): void => {
        if (!editor || disabled) {
            return;
        }

        command();
    };

    const insertImageWithConditionalWidth = async (source: string): Promise<void> => {
        if (!editor) {
            return;
        }

        const width = await new Promise<number>((resolve) => {
            const image = new window.Image();

            image.onload = () => resolve(image.naturalWidth || 0);
            image.onerror = () => resolve(0);
            image.src = source;
        });

        if (width > 800) {
            editor.chain().focus().setImage({ src: source, width: 800 }).run();
            return;
        }

        editor.chain().focus().setImage({ src: source }).run();
    };

    const insertImageFromPicker = async (filePickerResult: IFilePickerResult): Promise<void> => {
        if (!editor) {
            return;
        }

        const directSource = filePickerResult.fileAbsoluteUrl || filePickerResult.previewDataUrl || '';

        if (directSource) {
            await insertImageWithConditionalWidth(directSource);
            return;
        }

        const fileContent = await filePickerResult.downloadFileContent();
        const fileReader = new FileReader();

        await new Promise<void>((resolve, reject) => {
            fileReader.onload = () => {
                if (typeof fileReader.result === 'string') {
                    insertImageWithConditionalWidth(fileReader.result)
                        .then(() => resolve())
                        .catch(reject);
                    return;
                }

                reject(new Error('Unable to read file content as data URL.'));
            };

            fileReader.onerror = () => reject(fileReader.error || new Error('File read failed.'));
            fileReader.readAsDataURL(fileContent);
        });
    };

    const handlePickerSave = async (selectedFiles: IFilePickerResult[]): Promise<void> => {
        setIsFilePickerOpen(false);

        if (!selectedFiles.length) {
            return;
        }

        await insertImageFromPicker(selectedFiles[0]);
    };

    const buildGoodsResultsTableHtml = (): string => {
        const headerHtml = visibleGoodsColumnKeys
            .map((key) => {
                const title = key === goodsCommentColumnKey
                    ? goodsCommentColumnTitle
                    : (goodsColumnTitleByKey[key] || key);
                return `<th>${escapeHtml(title)}</th>`;
            })
            .join('');

        const rowsHtml = goodsSearchResults
            .map((item) => {
                const cellsHtml = visibleGoodsColumnKeys
                    .map((key) => {
                        const value = key === goodsCommentColumnKey ? '' : String(item[key] ?? '');
                        return `<td>${escapeHtml(value)}</td>`;
                    })
                    .join('');

                return `<tr>${cellsHtml}</tr>`;
            })
            .join('');

        return `<table><thead><tr>${headerHtml}</tr></thead><tbody>${rowsHtml}</tbody></table>`;
    };

    const buildGoodsResultsTableText = (): string => {
        const headerText = visibleGoodsColumnKeys
            .map((key) => key === goodsCommentColumnKey ? goodsCommentColumnTitle : (goodsColumnTitleByKey[key] || key))
            .join('\t');
        const rowsText = goodsSearchResults.map((item) =>
            visibleGoodsColumnKeys
                .map((key) => key === goodsCommentColumnKey ? '' : String(item[key] ?? ''))
                .join('\t')
        );

        return [headerText].concat(rowsText).join('\n');
    };

    const resetGoodsSearchState = (): void => {
        setGoodsSearchText('');
        setGoodsSearchResults([]);
        setGoodsSearchError('');
        setGoodsCopyMessage('');
        setHasSearchedGoods(false);
    };

    const runGoodsSearch = React.useCallback((eanValues: string[], selectedColumnKeys: string[]): void => {
        if (!goodsListId) {
            setGoodsSearchResults([]);
            setGoodsSearchError('A terméklista nincs beállítva a webpart tulajdonságai között.');
            setGoodsCopyMessage('');
            setIsSearchingGoods(false);
            return;
        }

        const requestId = ++latestGoodsSearchRequestIdRef.current;

        setIsSearchingGoods(true);
        setGoodsSearchError('');
        setGoodsCopyMessage('');
        setGoodsSearchResults([]);

        pnpListsServices.searchGoodsByEanList(goodsListId, eanValues, selectedColumnKeys).then((results) => {
            if (requestId !== latestGoodsSearchRequestIdRef.current) {
                return;
            }

            setGoodsSearchResults(results);
            setIsSearchingGoods(false);
        }).catch((error) => {
            if (requestId !== latestGoodsSearchRequestIdRef.current) {
                return;
            }

            console.error('Error searching goods:', error);
            setGoodsSearchResults([]);
            setGoodsSearchError('Hiba történt a cikkek keresése közben.');
            setGoodsCopyMessage('');
            setIsSearchingGoods(false);
        });
    }, [goodsListId]);

    React.useEffect(() => {
        if (!hasSearchedGoods) {
            return;
        }

        if (skipNextAutoRefreshRef.current) {
            skipNextAutoRefreshRef.current = false;
            return;
        }

        const eanValues = parseEansFromInput(goodsSearchText);
        if (!eanValues.length) {
            return;
        }

        runGoodsSearch(eanValues, effectiveSelectedGoodsColumnKeys);
    }, [effectiveSelectedGoodsColumnKeys, goodsSearchText, hasSearchedGoods, runGoodsSearch]);

    const handleCopyGoodsTable = async (): Promise<void> => {
        if (!goodsSearchResults.length || !visibleGoodsColumnKeys.length) {
            return;
        }

        const htmlTable = buildGoodsResultsTableHtml();
        const textTable = buildGoodsResultsTableText();

        const eansFromInput = parseEansFromInput(goodsSearchText);
        const eanResultColumn = visibleGoodsColumnKeys.find((key) => {
            if (key === goodsCommentColumnKey) {
                return false;
            }

            const keyToken = toHeaderToken(key);
            const titleToken = toHeaderToken(goodsColumnTitleByKey[key] || key);
            return keyToken.indexOf('ean') !== -1 || titleToken.indexOf('ean') !== -1;
        });
        const cikkszamResultColumn = visibleGoodsColumnKeys.find((key) => {
            if (key === goodsCommentColumnKey) {
                return false;
            }

            const keyToken = toHeaderToken(key);
            const titleToken = toHeaderToken(goodsColumnTitleByKey[key] || key);
            return keyToken.indexOf('cikkszam') !== -1 || titleToken.indexOf('cikkszam') !== -1;
        });
        const eansFromResults = eanResultColumn
            ? goodsSearchResults
                .map((item) => normalizeEanCandidate(String(item[eanResultColumn] ?? '')))
                .filter((item, index, items) => item.length > 0 && items.indexOf(item) === index)
            : [];
        const cikkszamFromResults = cikkszamResultColumn
            ? goodsSearchResults
                .map((item) => normalizeProductCodeCandidate(String(item[cikkszamResultColumn] ?? '')))
                .filter((item, index, items) => item.length > 0 && items.indexOf(item) === index)
            : [];

        const productCodesToSync = eansFromInput
            .concat(eansFromResults)
            .concat(cikkszamFromResults)
            .filter((item, index, items) => item.length > 0 && items.indexOf(item) === index);

        if (onEansDetected && productCodesToSync.length) {
            onEansDetected(productCodesToSync);
        }

        try {
            if (navigator.clipboard && typeof ClipboardItem !== 'undefined') {
                await navigator.clipboard.write([
                    new ClipboardItem({
                        'text/html': new Blob([htmlTable], { type: 'text/html' }),
                        'text/plain': new Blob([textTable], { type: 'text/plain' }),
                    }),
                ]);
            } else if (navigator.clipboard) {
                await navigator.clipboard.writeText(textTable);
            } else {
                throw new Error('Clipboard API is not available.');
            }

            setGoodsCopyMessage('A találati tábla a vágólapra került.');
            setGoodsSearchError('');
        } catch {
            setGoodsCopyMessage('');
            setGoodsSearchError('Nem sikerült a táblát a vágólapra másolni.');
        }
    };

    const handleSearchGoods = (): void => {
        setHasSearchedGoods(true);
        const eanValues = parseEansFromInput(goodsSearchText);

        if (!eanValues.length) {
            setGoodsSearchResults([]);
            setGoodsSearchError('Adj meg legalább egy EAN számot a kereséshez.');
            setGoodsCopyMessage('');
            return;
        }

        skipNextAutoRefreshRef.current = true;
        runGoodsSearch(eanValues, effectiveSelectedGoodsColumnKeys);
    };

    if (!editor) {
        return (
            <div className={styles.editorRoot}>
                <Label>{label}</Label>
            </div>
        );
    }

    return (
        <Stack tokens={{ childrenGap: 8 }} className={styles.editorRoot}>
            {showGoodsSearch ? (
                <Stack horizontal verticalAlign="center" horizontalAlign="space-between" tokens={{ childrenGap: 8 }}>
                    <Label>{label}</Label>
                    <DefaultButton text="Áru kereső" disabled={!!disabled} onClick={() => setIsGoodsSearchOpen(true)} />
                </Stack>
            ) : (
                <Label>{label}</Label>
            )}
            <div className={styles.toolbar}>
                <IconButton
                    iconProps={{ iconName: 'ClearFormatting' }}
                    className={styles.toolButton}
                    disabled={!!disabled}
                    onClick={() => runCommand(() => editor.chain().focus().clearNodes().unsetAllMarks().run())}
                />
                <IconButton
                    iconProps={{ iconName: 'Undo' }}
                    className={styles.toolButton}
                    disabled={!!disabled || !editor.can().undo()}
                    onClick={() => runCommand(() => editor.chain().focus().undo().run())}
                />
                <IconButton
                    iconProps={{ iconName: 'Redo' }}
                    className={styles.toolButton}
                    disabled={!!disabled || !editor.can().redo()}
                    onClick={() => runCommand(() => editor.chain().focus().redo().run())}
                />
                <IconButton
                    iconProps={{ iconName: 'Bold' }}
                    className={styles.toolButton}
                    disabled={!!disabled}
                    onClick={() => runCommand(() => editor.chain().focus().toggleBold().run())}
                />
                <IconButton
                    iconProps={{ iconName: 'Italic' }}
                    className={styles.toolButton}
                    disabled={!!disabled}
                    onClick={() => runCommand(() => editor.chain().focus().toggleItalic().run())}
                />
                <IconButton
                    iconProps={{ iconName: 'Strikethrough' }}
                    className={styles.toolButton}
                    disabled={!!disabled}
                    onClick={() => runCommand(() => editor.chain().focus().toggleStrike().run())}
                />
                <IconButton
                    iconProps={{ iconName: 'BulletedList' }}
                    className={styles.toolButton}
                    disabled={!!disabled}
                    onClick={() => runCommand(() => editor.chain().focus().toggleBulletList().run())}
                />
                <IconButton
                    iconProps={{ iconName: 'NumberedList' }}
                    className={styles.toolButton}
                    disabled={!!disabled}
                    onClick={() => runCommand(() => editor.chain().focus().toggleOrderedList().run())}
                />
                <IconButton
                    iconProps={{ iconName: 'Header1' }}
                    className={styles.toolButton}
                    disabled={!!disabled}
                    onClick={() => runCommand(() => editor.chain().focus().toggleHeading({ level: 1 }).run())}
                />
                <IconButton
                    iconProps={{ iconName: 'Header2' }}
                    className={styles.toolButton}
                    disabled={!!disabled}
                    onClick={() => runCommand(() => editor.chain().focus().toggleHeading({ level: 2 }).run())}
                />
                <IconButton
                    iconProps={{ iconName: 'Header3' }}
                    className={styles.toolButton}
                    disabled={!!disabled}
                    onClick={() => runCommand(() => editor.chain().focus().toggleHeading({ level: 3 }).run())}
                />
                <IconButton
                    iconProps={{ iconName: 'AlignLeft' }}
                    className={styles.toolButton}
                    disabled={!!disabled}
                    onClick={() => runCommand(() => editor.chain().focus().setTextAlign('left').run())}
                />
                <IconButton
                    iconProps={{ iconName: 'AlignCenter' }}
                    className={styles.toolButton}
                    disabled={!!disabled}
                    onClick={() => runCommand(() => editor.chain().focus().setTextAlign('center').run())}
                />
                <IconButton
                    iconProps={{ iconName: 'AlignRight' }}
                    className={styles.toolButton}
                    disabled={!!disabled}
                    onClick={() => runCommand(() => editor.chain().focus().setTextAlign('right').run())}
                />
                <IconButton
                    menuProps={{
                        items: [
                            {
                                key: 'insertTable',
                                text: 'Táblázat beszúrása',
                                onClick: () => runCommand(() => editor.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run())
                            },
                            {
                                key: 'addColumnBefore',
                                text: 'Oszlop beszúrása előtt',
                                onClick: () => runCommand(() => editor.chain().focus().addColumnBefore().run())
                            },
                            {
                                key: 'addColumnAfter',
                                text: 'Oszlop beszúrása után',
                                onClick: () => runCommand(() => editor.chain().focus().addColumnAfter().run())
                            },
                            {
                                key: 'deleteColumn',
                                text: 'Oszlop törlése',
                                onClick: () => runCommand(() => editor.chain().focus().deleteColumn().run())
                            },
                            {
                                key: 'addRowBefore',
                                text: 'Sor beszúrása előtt',
                                onClick: () => runCommand(() => editor.chain().focus().addRowBefore().run())
                            },
                            {
                                key: 'addRowAfter',
                                text: 'Sor beszúrása után',
                                onClick: () => runCommand(() => editor.chain().focus().addRowAfter().run())
                            },
                            {
                                key: 'deleteRow',
                                text: 'Sor törlése',
                                onClick: () => runCommand(() => editor.chain().focus().deleteRow().run())
                            },
                            {
                                key: 'deleteTable',
                                text: 'Táblázat törlése',
                                onClick: () => runCommand(() => editor.chain().focus().deleteTable().run())
                            },
                            {
                                key: 'mergeCells',
                                text: 'Cellák egyesítése',
                                onClick: () => runCommand(() => editor.chain().focus().mergeCells().run())
                            },
                            {
                                key: 'splitCell',
                                text: 'Cellák szétválasztása',
                                onClick: () => runCommand(() => editor.chain().focus().splitCell().run())
                            },
                            {
                                key: 'toggleHeaderRow',
                                text: 'Fejléc sor beállítása/törlése',
                                onClick: () => runCommand(() => editor.chain().focus().toggleHeaderRow().run())
                            },
                            {
                                key: 'toggleHeaderColumn',
                                text: 'Fejléc oszlop beállítása/törlése',
                                onClick: () => runCommand(() => editor.chain().focus().toggleHeaderColumn().run())
                            },
                            {
                                key: 'toggleHeaderCell',
                                text: 'Fejléc cella beállítása/törlése',
                                onClick: () => runCommand(() => editor.chain().focus().toggleHeaderCell().run())
                            }
                        ],
                        directionalHintFixed: true
                    }}
                    iconProps={{ iconName: 'Table' }}
                    title="Táblázat"
                    ariaLabel="Táblázat"
                    className={styles.toolButton}
                    disabled={!!disabled}
                />
                <IconButton
                    iconProps={{ iconName: 'Picture' }}
                    className={styles.toolButton}
                    disabled={!!disabled || !context}
                    onClick={() => setIsFilePickerOpen(true)}
                />
                <IconButton
                    iconProps={{ iconName: 'CalculatorSubtract' }}
                    className={styles.toolButton}
                    disabled={!!disabled}
                    onClick={() => runCommand(() => editor.chain().focus().setHorizontalRule().run())}
                />
            </div>
            <div className={styles.editorSurface}>
                <EditorContent editor={editor} />
            </div>
            {context && (
                <FilePicker
                    context={context as any}
                    hidden={true}
                    isPanelOpen={isFilePickerOpen}
                    onSave={handlePickerSave}
                    onCancel={() => setIsFilePickerOpen(false)}
                    accepts={['.jpg', '.jpeg', '.png', '.gif', '.webp', '.svg']}
                    hideWebSearchTab={true}
                    hideStockImages={true}
                    hideOrganisationalAssetTab={true}
                    hideOneDriveTab={true}
                    hideRecentTab={true}
                />
            )}
            <Modal
                isOpen={isGoodsSearchOpen}
                isBlocking={false}
                onDismiss={() => {
                    setIsGoodsSearchOpen(false);
                    resetGoodsSearchState();
                }}
                styles={{
                    main: {
                        width: '1100px',
                        minWidth: '1100px',
                        maxWidth: '1100px',
                    },
                }}
            >
                <Stack tokens={{ childrenGap: 12, padding: 20 }}>
                    <Stack horizontal verticalAlign="center" horizontalAlign="space-between">
                        <Text variant="xLarge">Cikk kereső</Text>
                        <IconButton
                            iconProps={{ iconName: 'Cancel' }}
                            title="Bezárás"
                            ariaLabel="Bezárás"
                            onClick={() => {
                                setIsGoodsSearchOpen(false);
                                resetGoodsSearchState();
                            }}
                        />
                    </Stack>
                    <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
                        <TextField
                            label="EAN számok"
                            placeholder="Adj meg egy vagy több EAN számot, soronként vagy vesszővel elválasztva"
                            value={goodsSearchText}
                            multiline
                            rows={4}
                            styles={{ root: { flexGrow: 1 } }}
                            onChange={(_, nextValue) => {
                                setGoodsSearchText(nextValue ?? '');
                                setHasSearchedGoods(false);
                            }}
                            onKeyDown={(event) => {
                                if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
                                    event.preventDefault();
                                    handleSearchGoods();
                                }
                            }}
                        />
                        <DefaultButton text="Keresés" onClick={handleSearchGoods} disabled={isSearchingGoods} />
                    </Stack>
                    <ComboBox
                        label="Táblázat oszlopok"
                        placeholder="Válassz oszlopokat"
                        multiSelect={true}
                        allowFreeform={false}
                        autoComplete="on"
                        useComboBoxAsMenuWidth={true}
                        disabled={isLoadingGoodsColumns || isSearchingGoods || goodsColumnOptions.length === 0}
                        options={goodsColumnOptions}
                        onChange={(_, option) => {
                            if (!option) {
                                return;
                            }

                            const optionKey = String(option.key);

                            setSelectedGoodsColumnKeys((previous) => {
                                const isCurrentlySelected = previous.indexOf(optionKey) !== -1;
                                const nextSelectedKeys = isCurrentlySelected
                                    ? previous.filter((key) => key !== optionKey)
                                    : previous.concat(optionKey);

                                if (hasSearchedGoods) {
                                    const eanValues = parseEansFromInput(goodsSearchText);
                                    if (eanValues.length > 0) {
                                        skipNextAutoRefreshRef.current = true;
                                        runGoodsSearch(eanValues, nextSelectedKeys);
                                    }
                                }

                                return nextSelectedKeys;
                            });
                        }}
                    />
                    {isLoadingGoodsColumns && (
                        <MessageBar messageBarType={MessageBarType.info}>Cikk oszlopok betöltése...</MessageBar>
                    )}
                    {goodsSearchError && (
                        <MessageBar messageBarType={MessageBarType.error}>{goodsSearchError}</MessageBar>
                    )}
                    {goodsCopyMessage && (
                        <MessageBar messageBarType={MessageBarType.success}>{goodsCopyMessage}</MessageBar>
                    )}
                    {isSearchingGoods && (
                        <MessageBar messageBarType={MessageBarType.info}>Keresés folyamatban...</MessageBar>
                    )}
                    {!isSearchingGoods && hasSearchedGoods && goodsSearchResults.length === 0 && !goodsSearchError && (
                        <MessageBar messageBarType={MessageBarType.warning}>Nincs találat a megadott EAN számokra.</MessageBar>
                    )}
                    {goodsSearchResults.length > 0 && visibleGoodsColumnKeys.length > 0 && (
                        <div style={{ maxHeight: 360, overflow: 'auto', border: '1px solid #edebe9', borderRadius: 8 }}>
                            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                                <thead>
                                    <tr>
                                        {visibleGoodsColumnKeys.map((key) => (
                                            <th key={key} style={{ textAlign: 'left', borderBottom: '1px solid #edebe9', padding: '8px 10px' }}>
                                                {key === goodsCommentColumnKey ? goodsCommentColumnTitle : (goodsColumnTitleByKey[key] || key)}
                                            </th>
                                        ))}
                                    </tr>
                                </thead>
                                <tbody>
                                    {goodsSearchResults.map((item, index) => (
                                        <tr key={item.Id ?? index}>
                                            {visibleGoodsColumnKeys.map((key) => (
                                                <td key={key} style={{ borderBottom: '1px solid #f3f2f1', padding: '7px 10px' }}>
                                                    {key === goodsCommentColumnKey ? '' : String(item[key] ?? '')}
                                                </td>
                                            ))}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                    <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
                        <DefaultButton text="Vágólapra másolás" onClick={handleCopyGoodsTable} disabled={!goodsSearchResults.length} />
                        <DefaultButton
                            text="Bezárás"
                            onClick={() => {
                                setIsGoodsSearchOpen(false);
                                resetGoodsSearchState();
                            }}
                        />
                    </Stack>
                </Stack>
            </Modal>
        </Stack>
    );
};
