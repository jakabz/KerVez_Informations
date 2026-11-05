import { spfi, SPFI, SPFx as spSPFx } from "@pnp/sp";
import { GraphFI, graphfi, SPFx as graphSPFx } from "@pnp/graph";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/batching";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";
import "@pnp/graph/users";
import "@pnp/graph/groups";
import "@pnp/graph/members";
import "@pnp/graph/directory-objects";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { IAkcio, IAkcioDetails, IAkcioListItem } from "../models/IAkcio";
import { IShopListItem } from "../models/IShop";
import { IRelatedFileItem } from "../models/IRelatedFile";

export class ListViewService {
    public sp: SPFI = spfi();
    public graph: GraphFI = graphfi();
    private statusFieldNameByListId: Record<string, string> = {};
    private actionTypeFieldNameByListId: Record<string, string> = {};
    private shopTypeFieldNameByListId: Record<string, string> = {};
    private readonly goodsSearchViewFields: string[] = ['ID', 'Title', 'Cikkszam', 'EANszam', 'Megnevezes', 'Ar', 'AkciosAr'];
    private readonly goodsResultFieldAliases: { [key: string]: string[] } = {
        Cikkszam: ['Cikkszam', 'Cikksz_x00e1_m', 'Cikkszam_x0020_', 'CikkSzam', 'ArticleNumber', 'Sku', 'SKU'],
        EANszam: ['EANszam', 'EanSzam', 'EAN', 'Ean', 'Vonalkod', 'Barcode'],
        Megnevezes: ['Megnevezes', 'Megnevez_x00e9_s', 'Name', 'TermekNev', 'Description'],
        Ar: ['Ar', 'Price', 'NetPrice', 'BruttoAr', 'Brutto_x00e1_r'],
        AkciosAr: ['AkciosAr', 'Akcios_x00e1_r', 'DiscountPrice', 'SalePrice', 'PromoPrice'],
    };
    private readonly goodsSearchDefaultSystemFields: string[] = [
        'id',
        'title',
        'cikkszam',
        'eanszam',
        'megnevezes',
        'ar',
        'akciosar',
    ];

    public setup(context: WebPartContext): void {
        this.sp = spfi().using(spSPFx(context));
        this.graph = graphfi().using(graphSPFx(context));
    }

    public async isMemberOfGroups(groups: IPropertyFieldGroupOrPerson[], userLoginName: string): Promise<boolean> {
        for (const group of groups) {
            if(!group || !group.id) {
                continue;
            }
            const groupId = group.id.indexOf('|') > -1 ? group.id.split('|')[group.id.split('|').length - 1] : group.id;
            const groupMembers = await this.graph.groups.getById(groupId).members();
            if (groupMembers.some(member => member.userPrincipalName === userLoginName)) {
                return true;
            }
        }
        return false;
    }

    /**
     * Creates the list item without setting its status, so that the caller already knows the new
     * item's id if setActionStatus fails and a retry can update the item instead of creating another one.
     */
    public async createAction(listId: string, data: IAkcio): Promise<number> {
        const itemData = await this.buildActionItemData(listId, data);

        const addResult = await this.sp.web.lists.getById(listId).items.add(itemData);
        const createdItemId = this.tryResolveAddedItemId(addResult);

        if (!createdItemId) {
            throw new Error('Az információ létrejött, de az azonosítója nem állapítható meg. Zárd be a panelt, és a listából nyisd meg az információt.');
        }

        return createdItemId;
    }

    private tryResolveAddedItemId(addResult: unknown): number | undefined {
        if (!addResult || typeof addResult !== 'object') {
            return undefined;
        }

        const asAny = addResult as any;
        const candidates = [
            asAny?.data?.Id,
            asAny?.data?.ID,
            asAny?.Id,
            asAny?.ID,
            asAny?.item?.Id,
            asAny?.item?.ID,
        ];

        for (const candidate of candidates) {
            const parsed = Number(candidate);
            if (!Number.isNaN(parsed) && parsed > 0) {
                return parsed;
            }
        }

        return undefined;
    }

    public async updateAction(listId: string, itemId: number, data: IAkcio): Promise<void> {
        const itemData = await this.buildActionItemData(listId, data);

        await this.sp.web.lists.getById(listId).items.getById(itemId).update(itemData);
        await this.setActionStatus(listId, itemId, data.status);
    }

    public async getActionById(listId: string, itemId: number): Promise<IAkcioDetails> {
        const statusFieldName = await this.getStatusFieldName(listId);
        const item = await this.sp.web.lists.getById(listId).items
            .getById(itemId)
            .select(
                'Id',
                'Title',
                'ActionType',
                statusFieldName,
                'StartDate',
                'EndDate',
                'PublishStartDate',
                'Description',
                'RelatedDocuments',
                'RelatedShopsType',
                'RelatedShops',
                'RelatedProducts',
                'Images',
                'ResponsibleEmail',
                'ResponsiblePerson/Title',
                'ResponsiblePerson/EMail',
                'ResponsiblePerson/Name'
            )
            .expand('ResponsiblePerson')();

        return {
            id: item.Id,
            title: item.Title ?? '',
            actionType: item.ActionType ?? '',
            status: item[statusFieldName] ?? 'Piszkozat',
            startDate: item.StartDate ? new Date(item.StartDate) : undefined,
            endDate: item.EndDate ? new Date(item.EndDate) : undefined,
            publishStartDate: item.PublishStartDate ? new Date(item.PublishStartDate) : undefined,
            description: item.Description ?? '',
            responsiblePersonLoginName: item.ResponsiblePerson?.Name ?? '',
            responsiblePersonTitle: item.ResponsiblePerson?.Title ?? '',
            responsiblePersonEmail: item.ResponsiblePerson?.EMail ?? '',
            responsibleEmail: item.ResponsibleEmail ?? '',
            relatedDocuments: item.RelatedDocuments ?? '',
            relatedShopsType: item.RelatedShopsType ?? 'all',
            relatedShops: item.RelatedShops ?? '',
            relatedProducts: item.RelatedProducts ?? '',
            images: item.Images ?? '',
        };
    }

    public async getActionTypeChoices(listId: string): Promise<string[]> {
        const actionTypeFieldName = await this.getActionTypeFieldName(listId);

        if (!actionTypeFieldName) {
            return [];
        }

        try {
            const field = await this.sp.web.lists
                .getById(listId)
                .fields
                .getByInternalNameOrTitle(actionTypeFieldName)
                .select('Choices')();

            const rawChoices = Array.isArray((field as any)?.Choices) ? (field as any).Choices : [];
            const choices = rawChoices
                .map((choice: unknown) => String(choice ?? '').trim())
                .filter((choice: string) => choice.length > 0);

            return Array.from(new Set(choices));
        } catch {
            return [];
        }
    }

    private async buildActionItemData(listId: string, data: IAkcio): Promise<Record<string, unknown>> {
        const statusFieldName = await this.getStatusFieldName(listId);
        const itemData: Record<string, unknown> = {
            Title: data.title,
            ActionType: data.actionType,
            StartDate: data.startDate?.toISOString() ?? null,
            EndDate: data.endDate?.toISOString() ?? null,
            PublishStartDate: data.publishStartDate?.toISOString() ?? null,
            Description: data.description || null,
            RelatedDocuments: data.relatedDocuments || null,
            RelatedShopsType: data.relatedShopsType,
            RelatedShops: data.relatedShops || null,
            RelatedProducts: data.relatedProducts || null,
            Images: data.images || null,
            ResponsibleEmail: data.responsibleEmail || null,
        };

        itemData[statusFieldName] = data.status;

        if (data.responsiblePersonLoginName) {
            const userResult = await this.sp.web.ensureUser(data.responsiblePersonLoginName);
            itemData.ResponsiblePersonId = userResult.Id;
        } else {
            itemData.ResponsiblePersonId = null;
        }

        return itemData;
    }

    public async getActions(listId: string): Promise<IAkcioListItem[]> {
        const statusFieldName = await this.getStatusFieldName(listId);
        const items = await this.sp.web.lists.getById(listId).items
            .select('Id', 'Title', 'ActionType', statusFieldName, 'StartDate', 'EndDate', 'PublishStartDate', 'ResponsiblePerson/Title', 'RelatedProducts')
            .expand('ResponsiblePerson')
            .orderBy('StartDate', false)();

        return items.map(item => ({
            id: item.Id,
            title: item.Title ?? '',
            actionType: item.ActionType ?? '',
            status: item[statusFieldName] ?? '',
            startDate: item.StartDate ?? null,
            endDate: item.EndDate ?? null,
            publishStartDate: item.PublishStartDate ?? null,
            responsiblePersonTitle: item.ResponsiblePerson?.Title ?? '',
            relatedProducts: item.RelatedProducts ?? '',
        }));
    }

    private async getStatusFieldName(listId: string): Promise<string> {
        if (this.statusFieldNameByListId[listId]) {
            return this.statusFieldNameByListId[listId];
        }

        const normalize = (value: string): string => value
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, '');

        try {
            const fields = await this.sp.web.lists.getById(listId).fields.select('InternalName', 'Title', 'EntityPropertyName')();
            const preferred = fields.find((field: any) => normalize(String(field.Title ?? '')) === 'status' || normalize(String(field.Title ?? '')) === 'statusz');
            const byInternal = fields.find((field: any) => String(field.InternalName ?? '') === 'Status');
            const byEntity = fields.find((field: any) => String(field.EntityPropertyName ?? '') === 'Status');

            const resolved = String(
                preferred?.InternalName
                || preferred?.EntityPropertyName
                || byInternal?.InternalName
                || byInternal?.EntityPropertyName
                || byEntity?.InternalName
                || byEntity?.EntityPropertyName
                || 'Status'
            );

            this.statusFieldNameByListId[listId] = resolved;
            return resolved;
        } catch {
            this.statusFieldNameByListId[listId] = 'Status';
            return 'Status';
        }
    }

    private async getActionTypeFieldName(listId: string): Promise<string> {
        if (this.actionTypeFieldNameByListId[listId]) {
            return this.actionTypeFieldNameByListId[listId];
        }

        const normalize = (value: string): string => value
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, '');

        try {
            const fields = await this.sp.web.lists.getById(listId).fields.select('InternalName', 'Title', 'EntityPropertyName', 'TypeAsString')();
            const candidates = fields.filter((field: any) => String(field.TypeAsString ?? '') === 'Choice' || String(field.TypeAsString ?? '') === 'MultiChoice');

            const preferred = candidates.find((field: any) => {
                const title = normalize(String(field.Title ?? ''));
                const internal = normalize(String(field.InternalName ?? ''));
                const entity = normalize(String(field.EntityPropertyName ?? ''));

                return title === 'actiontype'
                    || title === 'akciotipusa'
                    || title === 'akciotipus'
                    || title === 'informaciotipusa'
                    || title === 'informaciotipus'
                    || internal === 'actiontype'
                    || entity === 'actiontype';
            });

            const resolved = String(
                preferred?.InternalName
                || preferred?.EntityPropertyName
                || ''
            );

            this.actionTypeFieldNameByListId[listId] = resolved;
            return resolved;
        } catch {
            this.actionTypeFieldNameByListId[listId] = 'ActionType';
            return 'ActionType';
        }
    }

    public async setActionStatus(listId: string, itemId: number, statusValue: string): Promise<void> {
        if (!statusValue) {
            return;
        }

        const resolvedStatusFieldName = await this.getStatusFieldName(listId);
        const fieldNameCandidates = Array.from(new Set([
            resolvedStatusFieldName,
            'Status',
            'St_x00e1_tusz',
            'OData__Status',
            'OData__St_x00e1_tusz',
            'Státusz',
        ].filter((value) => !!value)));

        let lastError: unknown;
        const attempts: string[] = [];

        for (const fieldName of fieldNameCandidates) {
            try {
                const result = await this.sp.web.lists.getById(listId).items.getById(itemId).validateUpdateListItem([
                    { FieldName: fieldName, FieldValue: statusValue },
                ]);

                const hasException = Array.isArray(result)
                    && result.some((entry: any) => !!entry?.HasException);

                if (!hasException) {
                    return;
                }

                attempts.push(`${fieldName}:has-exception`);
                lastError = result;
            } catch (error) {
                attempts.push(`${fieldName}:error`);
                lastError = error;
            }
        }

        if (lastError) {
            const details = (() => {
                if (lastError instanceof Error) {
                    return lastError.message;
                }

                try {
                    return JSON.stringify(lastError);
                } catch {
                    return String(lastError);
                }
            })();

            throw new Error(`Status update failed for list ${listId}, item ${itemId}. Attempts: ${attempts.join(', ')}. Details: ${details}`);
        }
    }

    public async getShops(listId: string): Promise<IShopListItem[]> {
        const items = await this.sp.web.lists.getById(listId).items
            .select('Id', 'Title')
            .orderBy('Title', true)
            .top(5000)();

        return items.map(item => ({
            id: item.Id,
            title: item.Title ?? '',
        }));
    }

    public async getShopTypeChoices(listId: string): Promise<string[]> {
        const shopTypeFieldName = await this.getShopTypeFieldName(listId);

        if (!shopTypeFieldName) {
            return [];
        }

        try {
            const field = await this.sp.web.lists
                .getById(listId)
                .fields
                .getByInternalNameOrTitle(shopTypeFieldName)
                .select('Choices')();

            const rawChoices = Array.isArray((field as any)?.Choices) ? (field as any).Choices : [];
            const choices = rawChoices
                .map((choice: unknown) => String(choice ?? '').trim())
                .filter((choice: string) => choice.length > 0);

            return Array.from(new Set(choices));
        } catch {
            return [];
        }
    }

    private async getShopTypeFieldName(listId: string): Promise<string> {
        if (this.shopTypeFieldNameByListId[listId]) {
            return this.shopTypeFieldNameByListId[listId];
        }

        const normalize = (value: string): string => value
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, '');

        try {
            const fields = await this.sp.web.lists.getById(listId).fields.select('InternalName', 'Title', 'EntityPropertyName', 'TypeAsString')();
            const candidates = fields.filter((field: any) => String(field.TypeAsString ?? '') === 'Choice' || String(field.TypeAsString ?? '') === 'MultiChoice');

            const preferred = candidates.find((field: any) => {
                const title = normalize(String(field.Title ?? ''));
                const internal = normalize(String(field.InternalName ?? ''));
                const entity = normalize(String(field.EntityPropertyName ?? ''));

                return title === 'shoptype'
                    || title === 'bolttipus'
                    || internal === 'shoptype'
                    || entity === 'shoptype';
            });

            const resolved = String(
                preferred?.InternalName
                || preferred?.EntityPropertyName
                || ''
            );

            this.shopTypeFieldNameByListId[listId] = resolved;
            return resolved;
        } catch {
            this.shopTypeFieldNameByListId[listId] = 'ShopType';
            return 'ShopType';
        }
    }

    public async getGoodsSearchColumns(goodsListId: string): Promise<Array<{ internalName: string; title: string }>> {
        if (!goodsListId) {
            return [];
        }

        const fields = await this.sp.web.lists.getById(goodsListId).fields.select(
            'InternalName',
            'Title',
            'TypeAsString',
            'Hidden',
            'ReadOnlyField'
        )();

        const allowedTypeTokens = [
            'text',
            'note',
            'number',
            'currency',
            'boolean',
            'datetime',
            'choice'
        ];

        const normalizedDefaultOrder = this.goodsSearchDefaultSystemFields;
        const selectableFields = fields
            .filter((field: any) => !field.Hidden)
            .filter((field: any) => !field.ReadOnlyField)
            .filter((field: any) => {
                const fieldType = String(field?.TypeAsString || '').toLowerCase();
                return allowedTypeTokens.indexOf(fieldType) !== -1;
            })
            .map((field: any) => ({
                internalName: String(field?.InternalName || '').trim(),
                title: String(field?.Title || field?.InternalName || '').trim(),
            }))
            .filter((field) => field.internalName !== '' && field.title !== '');

        const prioritizedFields = selectableFields
            .filter((field) => normalizedDefaultOrder.indexOf(this.normalizeFieldToken(field.internalName)) !== -1)
            .sort((a, b) => (
                normalizedDefaultOrder.indexOf(this.normalizeFieldToken(a.internalName))
                - normalizedDefaultOrder.indexOf(this.normalizeFieldToken(b.internalName))
            ));

        const remainingFields = selectableFields
            .filter((field) => normalizedDefaultOrder.indexOf(this.normalizeFieldToken(field.internalName)) === -1)
            .sort((a, b) => a.title.localeCompare(b.title, 'hu'));

        return prioritizedFields.concat(remainingFields);
    }

    public async searchGoodsByEanList(goodsListId: string, eanValues: string[], selectedColumnKeys: string[] = []): Promise<any[]> {
        if (!goodsListId || !eanValues.length) {
            return [];
        }

        const normalizedSearchTexts = eanValues
            .map((value) => value.trim())
            .filter((value, index, values) => value.length > 0 && values.indexOf(value) === index);

        if (!normalizedSearchTexts.length) {
            return [];
        }

        const goodsList = this.sp.web.lists.getById(goodsListId);
        const fields = await goodsList.fields.select('InternalName', 'Title', 'TypeAsString', 'Hidden')();
        const normalizedFields = fields.filter((field: any) => !field.Hidden);

        const eanFieldName = this.resolveFieldInternalName(normalizedFields, this.goodsResultFieldAliases.EANszam, 'EANszam');
        const eanFieldMeta = this.findFieldMeta(normalizedFields, eanFieldName);
        const eanFieldType = this.toCamlValueType(eanFieldMeta ? eanFieldMeta.TypeAsString : 'Text');

        const selectedFieldByAlias: { [key: string]: string } = {
            Cikkszam: this.resolveFieldInternalName(normalizedFields, this.goodsResultFieldAliases.Cikkszam, 'Cikkszam'),
            EANszam: eanFieldName,
            Megnevezes: this.resolveFieldInternalName(normalizedFields, this.goodsResultFieldAliases.Megnevezes, 'Megnevezes'),
            Ar: this.resolveFieldInternalName(normalizedFields, this.goodsResultFieldAliases.Ar, 'Ar'),
            AkciosAr: this.resolveFieldInternalName(normalizedFields, this.goodsResultFieldAliases.AkciosAr, 'AkciosAr')
        };

        const normalizedSelectedFields = selectedColumnKeys
            .map((fieldName) => String(fieldName || '').trim())
            .filter((fieldName, index, values) => fieldName !== '' && values.indexOf(fieldName) === index);

        const availableFieldNames = normalizedFields
            .map((field: any) => String(field?.InternalName || '').trim())
            .filter((fieldName) => fieldName !== '');

        const effectiveSelectedFieldNames = normalizedSelectedFields
            .filter((fieldName) => availableFieldNames.indexOf(fieldName) !== -1);

        const hasDynamicSelection = effectiveSelectedFieldNames.length > 0;

        const requestedViewFields: string[] = ['ID', 'Title'];
        if (hasDynamicSelection) {
            effectiveSelectedFieldNames.forEach((internalName) => {
                if (requestedViewFields.indexOf(internalName) === -1) {
                    requestedViewFields.push(internalName);
                }
            });
            if (requestedViewFields.indexOf(eanFieldName) === -1) {
                requestedViewFields.push(eanFieldName);
            }
        } else {
            Object.keys(selectedFieldByAlias).forEach((aliasKey) => {
                const internalName = selectedFieldByAlias[aliasKey];
                if (requestedViewFields.indexOf(internalName) === -1) {
                    requestedViewFields.push(internalName);
                }
            });
        }

        const effectiveViewFields = requestedViewFields.length > 2 ? requestedViewFields : this.goodsSearchViewFields;
        const results: any[] = [];
        const resultIds: { [key: number]: boolean } = {};
        const chunkSize = 20;

        const eanValueVariantsByInput: { [key: string]: string[] } = {};
        normalizedSearchTexts.forEach((value) => {
            const variants: string[] = [value];
            const numericOnly = value.replace(/\D+/g, '');
            if (numericOnly && variants.indexOf(numericOnly) === -1) {
                variants.push(numericOnly);
            }
            eanValueVariantsByInput[value] = variants;
        });

        for (let index = 0; index < normalizedSearchTexts.length; index += chunkSize) {
            const chunk = normalizedSearchTexts.slice(index, index + chunkSize);
            const chunkResults = await Promise.all(chunk.map(async (ean) => {
                const eanVariants = eanValueVariantsByInput[ean] || [ean];
                const searchAttempts: Array<{ value: string; valueType: 'Text' | 'Number' }> = [];

                eanVariants.forEach((variant) => {
                    searchAttempts.push({ value: variant, valueType: eanFieldType });
                    if (eanFieldType !== 'Text') {
                        searchAttempts.push({ value: variant, valueType: 'Text' });
                    }
                });

                for (let attemptIndex = 0; attemptIndex < searchAttempts.length; attemptIndex++) {
                    const currentAttempt = searchAttempts[attemptIndex];
                    if (currentAttempt.valueType === 'Number' && !/^\d+(\.\d+)?$/.test(currentAttempt.value)) {
                        continue;
                    }

                    const camlQuery = {
                        ViewXml: `
                            <View>
                                <ViewFields>
                                    ${effectiveViewFields.map((fieldName) => `<FieldRef Name='${fieldName}' />`).join('')}
                                </ViewFields>
                                <Query>
                                    <Where>
                                        <Eq>
                                            <FieldRef Name='${this.escapeCamlValue(eanFieldName)}' />
                                            <Value Type='${currentAttempt.valueType}'>${this.escapeCamlValue(currentAttempt.value)}</Value>
                                        </Eq>
                                    </Where>
                                </Query>
                                <RowLimit>100</RowLimit>
                            </View>`
                    };

                    const queryResult = await goodsList.getItemsByCAMLQuery(camlQuery);
                    if (queryResult.length > 0) {
                        return queryResult;
                    }
                }

                return [];
            }));

            chunkResults.forEach((items) => {
                items.forEach((item: any) => {
                    if (!resultIds[item.Id]) {
                        resultIds[item.Id] = true;
                        results.push(
                            hasDynamicSelection
                                ? this.mapDynamicGoodsResultItem(item, effectiveSelectedFieldNames)
                                : this.mapGoodsResultItem(item, selectedFieldByAlias)
                        );
                    }
                });
            });
        }

        if (!results.length) {
            for (let index = 0; index < normalizedSearchTexts.length; index++) {
                const value = normalizedSearchTexts[index];
                const numericOnly = value.replace(/\D+/g, '');
                if (!numericOnly) {
                    continue;
                }

                const camlQuery = {
                    ViewXml: `
                        <View>
                            <ViewFields>
                                ${effectiveViewFields.map((fieldName) => `<FieldRef Name='${fieldName}' />`).join('')}
                            </ViewFields>
                            <Query>
                                <Where>
                                    <BeginsWith>
                                        <FieldRef Name='${this.escapeCamlValue(eanFieldName)}' />
                                        <Value Type='Text'>${this.escapeCamlValue(numericOnly)}</Value>
                                    </BeginsWith>
                                </Where>
                            </Query>
                            <RowLimit>50</RowLimit>
                        </View>`
                };

                const fallbackItems = await goodsList.getItemsByCAMLQuery(camlQuery);
                fallbackItems.forEach((item: any) => {
                    if (!resultIds[item.Id]) {
                        resultIds[item.Id] = true;
                        results.push(
                            hasDynamicSelection
                                ? this.mapDynamicGoodsResultItem(item, effectiveSelectedFieldNames)
                                : this.mapGoodsResultItem(item, selectedFieldByAlias)
                        );
                    }
                });

                if (results.length > 0) {
                    break;
                }
            }
        }

        return results;
    }

    private normalizeFieldToken(value: string): string {
        return (value || '')
            .toLowerCase()
            .replace(/[^a-z0-9]/g, '');
    }

    private resolveFieldInternalName(fields: any[], aliases: string[], defaultName: string): string {
        const normalizedAliases = aliases.map((alias) => this.normalizeFieldToken(alias));

        for (let fieldIndex = 0; fieldIndex < fields.length; fieldIndex++) {
            const field = fields[fieldIndex];
            const internalToken = this.normalizeFieldToken(field.InternalName || '');
            const titleToken = this.normalizeFieldToken(field.Title || '');

            if (normalizedAliases.indexOf(internalToken) !== -1 || normalizedAliases.indexOf(titleToken) !== -1) {
                return field.InternalName;
            }
        }

        for (let fieldIndex = 0; fieldIndex < fields.length; fieldIndex++) {
            const field = fields[fieldIndex];
            const internalToken = this.normalizeFieldToken(field.InternalName || '');
            const titleToken = this.normalizeFieldToken(field.Title || '');

            for (let aliasIndex = 0; aliasIndex < normalizedAliases.length; aliasIndex++) {
                const aliasToken = normalizedAliases[aliasIndex];
                if (aliasToken.length < 3) {
                    continue;
                }

                if (
                    internalToken.indexOf(aliasToken) !== -1
                    || titleToken.indexOf(aliasToken) !== -1
                    || aliasToken.indexOf(internalToken) !== -1
                    || aliasToken.indexOf(titleToken) !== -1
                ) {
                    return field.InternalName;
                }
            }
        }

        return defaultName;
    }

    private findFieldMeta(fields: any[], internalName: string): any | null {
        for (let fieldIndex = 0; fieldIndex < fields.length; fieldIndex++) {
            if (fields[fieldIndex].InternalName === internalName) {
                return fields[fieldIndex];
            }
        }

        return null;
    }

    private toCamlValueType(fieldType: string): 'Text' | 'Number' {
        const normalizedType = (fieldType || '').toLowerCase();
        if (normalizedType === 'number' || normalizedType === 'currency') {
            return 'Number';
        }

        return 'Text';
    }

    private mapGoodsResultItem(item: any, aliasesToInternalNames: { [key: string]: string }): any {
        const mapped = {
            Id: item.Id,
            ID: item.ID,
            Title: item.Title
        } as any;

        Object.keys(aliasesToInternalNames).forEach((aliasKey) => {
            const internalName = aliasesToInternalNames[aliasKey];
            mapped[aliasKey] = item[aliasKey] !== undefined ? item[aliasKey] : item[internalName];
        });

        return mapped;
    }

    private mapDynamicGoodsResultItem(item: any, fieldInternalNames: string[]): any {
        const mapped = {
            Id: item.Id,
            ID: item.ID,
            Title: item.Title
        } as any;

        fieldInternalNames.forEach((fieldName) => {
            mapped[fieldName] = item[fieldName];
        });

        return mapped;
    }

    private escapeCamlValue(value: string): string {
        return value
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    private async getActionFolderServerRelativeUrl(libraryId: string, actionId: number): Promise<string> {
        const rootFolder = await this.sp.web.lists.getById(libraryId).rootFolder.select('ServerRelativeUrl')();
        return `${rootFolder.ServerRelativeUrl}/${actionId}`;
    }

    private async ensureActionFolderServerRelativeUrl(libraryId: string, actionId: number): Promise<string> {
        const folderServerRelativeUrl = await this.getActionFolderServerRelativeUrl(libraryId, actionId);

        try {
            await this.sp.web.folders.addUsingPath(folderServerRelativeUrl);
        } catch (error) {
            if (!this.isFolderAlreadyExistsError(error)) {
                throw error;
            }
        }

        return folderServerRelativeUrl;
    }

    private isFolderAlreadyExistsError(error: unknown): boolean {
        const raw = (() => {
            if (error instanceof Error) {
                return error.message;
            }

            if (typeof error === 'string') {
                return error;
            }

            try {
                return JSON.stringify(error);
            } catch {
                return String(error);
            }
        })();

        const normalized = raw
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '');

        return normalized.includes('-2130575257')
            || normalized.includes('already exists')
            || normalized.includes('mar letezik')
            || normalized.includes('file or folder')
            || normalized.includes('fajl vagy mappa mar letezik');
    }

    public async uploadActionRelatedFile(libraryId: string, actionId: number, fileName: string, fileContent: Blob): Promise<void> {
        const folderServerRelativeUrl = await this.ensureActionFolderServerRelativeUrl(libraryId, actionId);
        await this.sp.web.getFolderByServerRelativePath(folderServerRelativeUrl).files.addUsingPath(fileName, fileContent, { Overwrite: true });
    }

    public async getActionRelatedFiles(libraryId: string, actionId: number): Promise<IRelatedFileItem[]> {
        const folderServerRelativeUrl = await this.getActionFolderServerRelativeUrl(libraryId, actionId);
        let files: any[] = [];

        try {
            files = await this.sp.web.getFolderByServerRelativePath(folderServerRelativeUrl).files
                .select('Name', 'ServerRelativeUrl', 'Length', 'TimeCreated')
                .orderBy('Name', true)();
        } catch (error) {
            if (!this.isFolderMissingError(error)) {
                throw error;
            }

            return [];
        }

        return files.map((file: any) => ({
            name: file.Name ?? '',
            serverRelativeUrl: file.ServerRelativeUrl ?? '',
            size: Number(file.Length ?? 0),
            timeCreated: file.TimeCreated,
        }));
    }

    public async deleteActionRelatedFile(libraryId: string, actionId: number, fileServerRelativeUrl: string): Promise<void> {
        if (!libraryId || !actionId || !fileServerRelativeUrl) {
            return;
        }

        const folderServerRelativeUrl = await this.getActionFolderServerRelativeUrl(libraryId, actionId);
        const normalizedFolder = folderServerRelativeUrl.replace(/\/+$/, '');
        const normalizedFile = fileServerRelativeUrl.trim();

        if (!normalizedFile.startsWith(`${normalizedFolder}/`)) {
            throw new Error('A kijelölt fájl nem ehhez az információhoz tartozik.');
        }

        await this.sp.web.getFileByServerRelativePath(normalizedFile).delete();
    }

    private isFolderMissingError(error: unknown): boolean {
        const raw = (() => {
            if (error instanceof Error) {
                return error.message;
            }

            if (typeof error === 'string') {
                return error;
            }

            try {
                return JSON.stringify(error);
            } catch {
                return String(error);
            }
        })();

        const normalized = raw
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '');

        return normalized.includes('does not exist')
            || normalized.includes('cannot find')
            || normalized.includes('file not found')
            || normalized.includes('404')
            || normalized.includes('nem talalhato')
            || normalized.includes('nem letezik');
    }

}

const pnpListsServices = new ListViewService();
export default pnpListsServices;