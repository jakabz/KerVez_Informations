import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneLabel, PropertyPaneTextField, PropertyPaneToggle, type IPropertyPaneConfiguration, type IPropertyPaneField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { Akciok, IAkciokProps } from './components/Akciok';
import pnpListsServices from './services/pnpListServices';
import { IPropertyFieldGroupOrPerson, PrincipalType, PropertyFieldPeoplePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

export interface IAkciokWebPartProps {
  description: string;
  actionsListId: string;
  shopsListId: string;
  goodsListId: string;
  relatedDocumentLibraryId: string;
  tasksListPage: string;
  editorGroup: IPropertyFieldGroupOrPerson[];
  isDashboardView: boolean;
  actionTypeColors: Record<string, string>;
}

export default class AkciokWebPart extends BaseClientSideWebPart<IAkciokWebPartProps> {
  private actionTypeChoices: string[] = [];
  private isLoadingActionTypeChoices: boolean = false;
  private actionTypeChoicesError: string = '';

  private getDecodedActionTypeColors(): Record<string, string> {
    const source = this.properties.actionTypeColors || {};
    const decoded: Record<string, string> = {};

    Object.keys(source).forEach((encodedKey) => {
      const safeKey = (() => {
        try {
          return decodeURIComponent(encodedKey);
        } catch {
          return encodedKey;
        }
      })();

      decoded[safeKey] = source[encodedKey];
    });

    return decoded;
  }

  public render(): void {
    const actionTypeColors = this.getDecodedActionTypeColors();

    const element: React.ReactElement<IAkciokProps> = React.createElement(
      Akciok,
      {
        context: this.context,
        actionsListId: this.properties.actionsListId,
        shopsListId: this.properties.shopsListId,
        goodsListId: this.properties.goodsListId,
        relatedDocumentLibraryId: this.properties.relatedDocumentLibraryId,
        editorGroup: this.properties.editorGroup,
        tasksListPage: this.properties.tasksListPage,
        isDashboardView: !!this.properties.isDashboardView,
        actionTypeColors,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      pnpListsServices.setup(this.context);

      if (!this.properties.actionTypeColors) {
        this.properties.actionTypeColors = {};
      }
    });
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    await this.loadActionTypeChoices();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'actionsListId' && oldValue !== newValue) {
      void this.loadActionTypeChoices(true);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private getActionTypeColorKey(actionType: string): string {
    return encodeURIComponent(actionType);
  }

  private getActionTypeColor(actionType: string): string {
    const key = this.getActionTypeColorKey(actionType);
    return this.properties.actionTypeColors?.[key] || '#0078d4';
  }

  private async loadActionTypeChoices(forceRefresh: boolean = false): Promise<void> {
    if (this.isLoadingActionTypeChoices) {
      return;
    }

    if (!this.properties.actionsListId) {
      this.actionTypeChoices = [];
      this.actionTypeChoicesError = '';
      this.context.propertyPane.refresh();
      return;
    }

    if (!forceRefresh && this.actionTypeChoices.length > 0) {
      return;
    }

    this.isLoadingActionTypeChoices = true;
    this.actionTypeChoicesError = '';
    this.context.propertyPane.refresh();

    try {
      this.actionTypeChoices = await pnpListsServices.getActionTypeChoices(this.properties.actionsListId);
      this.actionTypeChoicesError = '';
    } catch (error) {
      this.actionTypeChoices = [];
      this.actionTypeChoicesError = error instanceof Error ? error.message : String(error);
    } finally {
      this.isLoadingActionTypeChoices = false;
      this.context.propertyPane.refresh();
    }
  }

  private getActionTypeColorGroupFields(): IPropertyPaneField<any>[] {
    if (!this.properties.actionsListId) {
      return [
        PropertyPaneLabel('actionTypeColorHelp', {
          text: 'Először válaszd ki az Információk listát, utána itt megjelennek az információ típusok.',
        }),
      ];
    }

    if (this.isLoadingActionTypeChoices) {
      return [
        PropertyPaneLabel('actionTypeColorLoading', {
          text: 'Információ típusok betöltése folyamatban...',
        }),
      ];
    }

    if (this.actionTypeChoicesError) {
      return [
        PropertyPaneLabel('actionTypeColorError', {
          text: `Nem sikerült betölteni az információ típusokat: ${this.actionTypeChoicesError}`,
        }),
      ];
    }

    if (!this.actionTypeChoices.length) {
      return [
        PropertyPaneLabel('actionTypeColorEmpty', {
          text: 'Az ActionType oszlopban nincs beállított választható érték.',
        }),
      ];
    }

    return this.actionTypeChoices.map((actionType) => {
      const key = this.getActionTypeColorKey(actionType);

      return PropertyFieldColorPicker(`actionTypeColors.${key}`, {
        label: actionType,
        selectedColor: this.getActionTypeColor(actionType),
        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
        properties: this.properties,
        disabled: false,
        isHidden: false,
        alphaSliderHidden: true,
        style: PropertyFieldColorPickerStyle.Inline,
        iconName: 'Color',
        key: `actionTypeColor_${key}`,
      });
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          groups: [
            // {
            //   groupName: 'Megjelenés',
            //   isCollapsed: true,
            //   groupFields: [
            //     PropertyPaneToggle('isDashboardView', {
            //       label: 'Nézet típusa',
            //       onText: 'Dashboard nézet',
            //       offText: 'Teljes nézet',
            //     })
            //   ]
            // },
            {
              groupName: 'Beállítások',
              isCollapsed: true,
              groupFields: [
                PropertyFieldListPicker('actionsListId', {
                  label: 'Információk lista',
                  selectedList: this.properties.actionsListId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: 'actionslistPickerFieldId'
                }),
                PropertyFieldListPicker('shopsListId', {
                  label: 'Boltok lista',
                  selectedList: this.properties.shopsListId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: 'shopslistPickerFieldId'
                }),
                PropertyFieldListPicker('goodsListId', {
                  label: 'Termékek lista',
                  selectedList: this.properties.goodsListId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: 'goodslistPickerFieldId'
                }),
                PropertyFieldListPicker('relatedDocumentLibraryId', {
                  label: 'Kapcsolódó dokumentumok könyvtára',
                  selectedList: this.properties.relatedDocumentLibraryId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: 'relatedDocumentLibraryPickerFieldId'
                }),
                PropertyPaneTextField('tasksListPage', {
                  label: 'Feladatok oldala',
                  value: this.properties.tasksListPage,
                  placeholder: '/sites/.../SitePages/...',
                }),
                PropertyFieldPeoplePicker('editorGroup', {
                  label: 'Szerkesztő csoport',
                  initialData: this.properties.editorGroup,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context as any,
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'editorGroupPeopleFieldId',
                  searchTextLimit: 3
                })
              ]
            },
            {
              groupName: 'Információ típus színek',
              isCollapsed: true,
              groupFields: this.getActionTypeColorGroupFields(),
            }
          ]
        }
      ]
    };
  }
}
