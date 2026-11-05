import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import * as React from 'react';
import { createGlobalStore, GlobalStoreProvider, useGlobalState } from '../store/GlobalStore';
import styles from './Akciok.module.scss';
import { Stack } from '@fluentui/react';
import { AkciokListHeader } from './AkciokListHeader';
import { AkciokList } from './AkciokList';
import { AkcioPanel } from './AkcioPanel';

export interface IAkciokProps {
  context: WebPartContext;
  actionsListId: string;
  shopsListId: string;
  goodsListId: string;
  relatedDocumentLibraryId: string;
  editorGroup: IPropertyFieldGroupOrPerson[];
  tasksListPage: string;
  isDashboardView: boolean;
  actionTypeColors: Record<string, string>;
}

const AkciokContent: React.FunctionComponent<IAkciokProps> = (props: React.PropsWithChildren<IAkciokProps>) => {

  const { 
    setContext,
    setActionsListId,
    setShopsListId,
    setGoodsListId,
    setEditorGroup,
    setRelatedDocumentLibraryId,
    setTasksListPage,
    setViewMode,
    resetState,
  } = useGlobalState();

  React.useEffect(() => {
    setContext(props.context);
    setActionsListId(props.actionsListId);
    setShopsListId(props.shopsListId);
    setGoodsListId(props.goodsListId);
    setRelatedDocumentLibraryId(props.relatedDocumentLibraryId);
    setEditorGroup(props.editorGroup);
    setTasksListPage(props.tasksListPage);
    setViewMode(props.isDashboardView ? 'dashboard' : 'full');
  }, [props.context, props.actionsListId, props.shopsListId, props.goodsListId, props.relatedDocumentLibraryId, props.editorGroup, props.tasksListPage, props.isDashboardView]);

  React.useEffect(() => {
    return () => {
      resetState();
    };
  }, [resetState]);

  return (
    <Stack className={styles.akciok}>
      <AkciokListHeader />
      <AkciokList actionTypeColors={props.actionTypeColors} />
      <AkcioPanel context={props.context} />
    </Stack>
  );
};

export const Akciok: React.FunctionComponent<IAkciokProps> = (props: React.PropsWithChildren<IAkciokProps>) => {
  const storeRef = React.useRef<ReturnType<typeof createGlobalStore> | null>(null);

  if (!storeRef.current) {
    storeRef.current = createGlobalStore();
  }

  return (
    <GlobalStoreProvider store={storeRef.current}>
      <AkciokContent {...props} />
    </GlobalStoreProvider>
  );
};