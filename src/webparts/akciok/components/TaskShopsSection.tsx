import * as React from 'react';
import {
    Pivot,
    PivotItem,
    Checkbox,
    SearchBox,
    List,
    Stack,
    Text,
} from '@fluentui/react';
import { useGlobalState } from '../store/GlobalStore';
import styles from './Akciok.module.scss';

export type ShopsSelectionType = 'all' | 'group' | 'selected';

export interface ITaskShopsSectionProps {
    shopsSelectionType: ShopsSelectionType;
    selectedShops: string[];
    onFormChange: (updates: Partial<{ shopsSelectionType: ShopsSelectionType; selectedShops: string[] }>) => void;
    disabled?: boolean;
}

export const TaskShopsSection: React.FunctionComponent<ITaskShopsSectionProps> = (props) => {
    const { shopListItems, shopTypeChoices } = useGlobalState();
    const [shopSearch, setShopSearch] = React.useState<string>('');
    const isDisabled = !!props.disabled;
    const groupedShopOptions = shopTypeChoices;

    return (
        <Stack className={styles.contentBox} tokens={{ childrenGap: 10 }}>
            <Text className={styles.sectionTitle}>Érintett boltok</Text>
            <Pivot
                linkFormat="tabs"
                selectedKey={props.shopsSelectionType}
                onLinkClick={(item) => {
                    if (isDisabled) {
                        return;
                    }

                    const nextKey = item?.props.itemKey as ShopsSelectionType | undefined;

                    if (!nextKey) {
                        return;
                    }

                    props.onFormChange({
                        shopsSelectionType: nextKey,
                        selectedShops: [],
                    });
                    setShopSearch('');
                }}
            >
                <PivotItem headerText="Minden bolt" itemKey="all" />
                <PivotItem headerText="Csoportosított boltok" itemKey="group">
                    <Stack horizontal tokens={{ childrenGap: 10 }} wrap styles={{ root: { marginTop: 10 } }}>
                        {groupedShopOptions.map((groupCode) => (
                            <Checkbox
                                key={groupCode}
                                label={groupCode}
                                checked={props.selectedShops.indexOf(groupCode) !== -1}
                                onChange={(_, checked) => {
                                    const newSelectedShops = checked
                                        ? [...props.selectedShops, groupCode]
                                        : props.selectedShops.filter(shop => shop !== groupCode);
                                    props.onFormChange({ selectedShops: newSelectedShops });
                                }}
                                disabled={isDisabled}
                            />
                        ))}
                        {!groupedShopOptions.length && (
                            <Text className={styles.viewerValue}>A Shops listában nincs beállított ShopType választási lehetőség.</Text>
                        )}
                    </Stack>
                </PivotItem>
                <PivotItem headerText="Kiválasztott boltok" itemKey="selected">
                    <Stack styles={{ root: { marginTop: 10 } }} tokens={{ childrenGap: 8 }}>
                        <SearchBox
                            placeholder="Kiválasztott boltok keresése..."
                            value={shopSearch}
                            disabled={isDisabled}
                            onChange={(_, newValue) => setShopSearch(newValue ?? '')}
                            onClear={() => setShopSearch('')}
                        />
                        <List
                            items={shopListItems.filter(shop => shop.title.toLowerCase().indexOf(shopSearch.toLowerCase()) !== -1)}
                            onRenderCell={(item, index) => (
                                <div key={index} className={styles.selectedShopItem}>
                                    {item && (
                                        <Checkbox
                                            label={item.title}
                                            checked={props.selectedShops.indexOf(String(item.id)) !== -1}
                                            onChange={(_, checked) => {
                                                const newSelectedShops = checked
                                                    ? [...props.selectedShops, String(item.id)]
                                                    : props.selectedShops.filter(shop => shop !== String(item.id));
                                                props.onFormChange({ selectedShops: newSelectedShops });
                                            }}
                                            disabled={isDisabled}
                                        />
                                    )}
                                </div>
                            )}
                        />
                    </Stack>
                </PivotItem>
            </Pivot>
        </Stack>
    );
};
