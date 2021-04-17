import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneCheckbox,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';

import * as strings from 'GroupPeopleWebPartStrings';

import GroupPeople from './components/GroupPeople';
import GroupPeoplePlaceHolder from './components/GroupPeoplePlaceholder';
import { IGroupPeopleProps } from './components/IGroupPeopleProps';

import SPGroupService from './services/SPGroupService';
import GraphGroupService from './services/GraphGroupService';
import MockSPGroupService from './mocks/MockSPGroupService';

import { PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import IGraphGroupService from './models/IGraphGroupService'
import PeopleCard from './models/PeopleCard';
import ISPGroupService from './models/ISPGroupService';
import { ISiteGroupInfo } from './models/ISiteGroupInfo';
import { ISiteUserInfo } from './models/ISiteUserInfo';
import { Utils } from './GroupPeopleUtils';
import IADGroup from './models/IADGroup';
import IADGroupPeople from './models/IADGroupPeople';

export interface IGroupPeopleWebPartProps {
  SPGroups: string;
  Layout: string;
  CustomTitle: string;
  ToggleTitle: boolean;
  PictureSize: string;
  HideWebPart: boolean;
  PictureUrl: string;
  LineOne: string;
  LineTwo: string;
  LineThree: string;
  NumberOfUsers: string;
  ToggleService: boolean;
}

/** Groupe People WebPart
 * @class
 * @extends
 */
export default class GroupPeopleWebPart extends BaseClientSideWebPart<IGroupPeopleWebPartProps> {

  /** List of available SharePoint site groups
   * @private
   */
  private _spSiteGrps: ISiteGroupInfo[];

  /** List of thte members available into the selected SharePoint group
   * @private
   */
  private _spGrpUsers: Array<PeopleCard>;

  /** Partial update statement
   * Detect if the webpart must be render partially in accordance with somes properties pane
   * @private
   */
  private _partialUpdateRender: boolean = false;

  /** SharePoint Group Service
   * @private
   */
  private _spGrpSvc: ISPGroupService;

  /**
   * SharePoint selected group title
   * @private
   */
  private _grpTitle: string;

  private _graphService: IGraphGroupService;
  private _adGroup: IADGroup[];

  /** Init WebPart
   * @returns
   * @protected
   */
  protected onInit(): Promise<void> {
    this._graphService = new GraphGroupService(this.context);
    this._spGrpSvc = Environment.type == EnvironmentType.Local ? new MockSPGroupService(this.context.pageContext.site.absoluteUrl) : new SPGroupService(this.context.pageContext.site.absoluteUrl);
    
    if (DisplayMode.Edit == this.displayMode) { /* Get all SharePoint groups only in edit mode */
        this._spGrpSvc.fetchSPGroups().then((spGroups: Array<ISiteGroupInfo>) => {
            this._spSiteGrps = spGroups;
            if (this.properties.SPGroups) {
                if(this._graphService.mClient) {
                    this.postRender();
                } else {
                    this._graphService.initialize().then((sucess) => {
                        this.getAdGroups((res) => {
                            this.getAdGroupUsers(this.properties.SPGroups, () => {
                                this._grpTitle = res.find((x) => { return x.id ===  this.properties.SPGroups}).displayName;
                                this.postRender();
                            })
                        });
                    }, (err) => {
                        console.error("Service not initialized", err);
                    })
                }
            } else {
                if(this._graphService.mClient) {
                    this.render();
                } else {
                    this._graphService.initialize().then((sucess) => {
                        this.render();
                    }, (err) => {
                        console.error("Service not initialized", err);
                    })
                }
            }
        });
    } 
    if (DisplayMode.Read == this.displayMode && this.properties.SPGroups) {
        this._graphService.initialize().then((sucess) => {
            if(this.properties.ToggleService) {
                this.getAdGroups((res) => {
                    this.getAdGroupUsers(this.properties.SPGroups, () => {
                        this._grpTitle = res.find((x) => { return x.id ===  this.properties.SPGroups}).displayName;
                        this.postRender();
                    })
                })
            } else {
                this._spGrpSvc.getSPGroup(parseInt(this.properties.SPGroups)).then((grp: ISiteGroupInfo) => {
                    this._grpTitle = grp.Title;
                    this.postRender();
                });
            }
        }, (err) => {
            console.error("Service not initialized", err);
        })
    }
    return super.onInit();
  }

  /** Default Render
   * @public
   */
  public render(): void {
    if (!this.properties.SPGroups || null == this.properties.SPGroups) {
      const element: React.ReactElement<GroupPeoplePlaceHolder> = React.createElement(GroupPeoplePlaceHolder);
      ReactDom.render(element, this.domElement);
    } else if (!this._partialUpdateRender && this.properties.SPGroups) {
      if(this.properties.ToggleService) {
          if(typeof this.properties.SPGroups == 'string') {
              this.getAdGroupUsers(this.properties.SPGroups, () => {})
          } else {
              this.getAdGroups((res) => {})
          }
      } else {
        this.getUsersGroup();
      }
    }
    this._partialUpdateRender = false;
  }

  private getNumber(): number {
    if (this.properties.NumberOfUsers && this.properties.NumberOfUsers.length > 0 && !isNaN(this.properties.NumberOfUsers as any)) {
        return parseInt(this.properties.NumberOfUsers)
    } else {
        // setting Defult to properties value
        this.properties.NumberOfUsers = 8+'';
        return 8
    }
  }

  /** Render the compact users layouts
   * @private
   */
  private postRender() {
    if ((this._spGrpUsers && this._spGrpUsers.length > 0) || (this._spGrpUsers && this._spGrpUsers.length == 0 && (undefined == this.properties.HideWebPart || false == this.properties.HideWebPart || DisplayMode.Edit == this.displayMode))) {
        if(this.properties.ToggleService) {
            this._grpTitle = this._grpTitle ? this._grpTitle : (undefined !== this.properties.SPGroups && null != this._adGroup) ? this._adGroup.find(g => g.id == this.properties.SPGroups).displayName : '';
        } else {
            this._grpTitle = this._grpTitle ? this._grpTitle : (undefined !== this.properties.SPGroups && null != this._spSiteGrps) ? this._spSiteGrps.find(g => g.Id == parseInt(this.properties.SPGroups)).Title : '';
        }
      const num =  this.getNumber();
      const element: React.ReactElement<IGroupPeopleProps> = React.createElement(GroupPeople, {
        title: (this.properties.CustomTitle && this.properties.CustomTitle.length > 0) ? this.properties.CustomTitle : this._grpTitle,
        users: (this._spGrpUsers && this._spGrpUsers.length > 0) ? this._spGrpUsers.sort((a: PeopleCard, b: PeopleCard) => (a.lineOne > b.lineOne) ? 1 : ((b.lineOne > a.lineOne) ? -1 : 0)) : new Array,
        size: PersonaSize[this.properties.PictureSize],
        displayTitle: this.properties.ToggleTitle,
        hide: (DisplayMode.Read == this.displayMode && this.properties.HideWebPart) ? true : false,
        numberOfUser: num
      });
      ReactDom.render(element, this.domElement);
    } else {
      this.onDispose();
    }
  }

  /**
   * Get all users from the selected SharePoint group and then populate the people cards
   */
  private getUsersGroup() {
    this._spGrpSvc.fetchUsersGroup(parseInt(this.properties.SPGroups)).then((users: Array<ISiteUserInfo>) => {
      return users;
    }).then((u: Array<ISiteUserInfo>) => {
      this._spGrpUsers = new Array;
      if (null != u && u.length > 0) {
        this.populatePeopleCards(u);
      } else {
        this.postRender();
      }
    });
  }

    private getAdGroups(callback) {
        this._graphService.getADGroups().then((res: IADGroup[]) => {
            this._adGroup = res;
            callback(res);
        }, (err) => {
            console.log("Graph group error", err)
        })
    }

    private getAdGroupUsers(id: string, callback) {
        this._graphService.getADGroupPeoples(id).then((res) => {
            this._spGrpUsers = new Array();
            res.forEach((item: IADGroupPeople) => {
                this._spGrpUsers.push(new PeopleCard(item.mail, '', item.displayName, item.jobTitle))
            })
            callback();
        }, (err) => {
            console.error('Error in fetching user', err);
        })
    }

  /**
   * Get the user profiles and populate the list of People Cards
   * @param u List of users informations members of the selected SharePoint groups
   */
  private populatePeopleCards(u: ISiteUserInfo[]) {
    let uCount = 0;
    u.forEach(user => {
      this._spGrpSvc.getUserProfile(user.LoginName).then((r) => {
        try {
          if (null != r && undefined != r) { // Ensure at least one user profile was found
            this._spGrpUsers.push(new PeopleCard(
              r.find(props => props.Key == 'AccountName').Value,
              r.find(props => props.Key == this.properties.PictureUrl) ? r.find(props => props.Key == this.properties.PictureUrl).Value : '',
              r.find(props => props.Key == this.properties.LineOne) ? r.find(props => props.Key == this.properties.LineOne).Value : '',
              r.find(props => props.Key == this.properties.LineTwo) ? r.find(props => props.Key == this.properties.LineTwo).Value : '',
              r.find(props => props.Key == this.properties.LineThree) ? r.find(props => props.Key == this.properties.LineThree).Value : ''
            ));
          }
        } catch (e) { /*console.log(e);*/ }
        uCount++;
        // Once all profiles are parsed, start the render
        if (uCount == u.length) {
          this.postRender();
        }
      });
    });
  }

  /** On Dispose
   * @protected
   */
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /** Version
   * @returns
   * @protected
   */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /** Customize the behavior of property pane change
   * @param targetProperty 
   * @param newValue 
   * @protected
   */
  protected onPropertyPaneFieldChanged(targetProperty: string, oldValue: any, newValue: any) {
    if ('CustomTitle' == targetProperty || 'ToggleTitle' == targetProperty) {
      this._partialUpdateRender = true;
      this.postRender();
    } else if ('PictureSize' == targetProperty) {
      this._partialUpdateRender = true;
      this.postRender();
    } else if ('NumberOfUsers' == targetProperty) {
      this._partialUpdateRender = true;
      this.postRender();
    } else if ('ToggleService' == targetProperty) {
        if(oldValue !== newValue && newValue) {
            this.getAdGroups((result) => {
                this._partialUpdateRender = true;
                this.postRender();
            })
        } else {
            this._partialUpdateRender = true;
            this.postRender();
        }
    } else if ('SPGroups' == targetProperty) {
      if ( oldValue !== newValue ) {
        this.getAdGroupUsers(newValue, () => {
            this._partialUpdateRender = true;
            this.postRender();
        })
      }
    }
  }

  /** Property Pane Configuration
   * @returns
   * @property
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupFields: [
                PropertyPaneToggle('ToggleService', {
                  label: strings.ToggleServiceLabel,
                  offText: strings.ToggleServiceOffText,
                  onText: strings.ToggleServiceOnText,
                }),
                PropertyPaneDropdown('SPGroups', {
                  label: (!this.properties.ToggleService) ? strings.DropdownGroupLabel : strings.DropdownADGroupLabel,
                  options: (!this.properties.ToggleService) ? Utils.convertGrpToOptions(this._spSiteGrps) : Utils.convertADGrpToOptions(this._adGroup)
                }),
                PropertyPaneLabel('LabelSeparator', {
                  text: ' '
                }),
                PropertyPaneToggle('ToggleTitle', {
                  label: strings.ToggleTitleLabel
                }),
                PropertyPaneTextField('CustomTitle', {
                  label: strings.CustomTitleLabel,
                  description: strings.CustomTitleDescription,
                  disabled: !this.properties.ToggleTitle
                }),
                PropertyPaneTextField('NumberOfUsers', {
                  label: strings.NumberOfUsersLabel,
                  description: strings.NumberOfUsersDescription,
                }),
                PropertyPaneDropdown('PictureSize', {
                  label: strings.PictureSize,
                  options: Utils.enumSizesToOptions(),
                  selectedKey: Utils.enumSizesToOptions()[0].key // Select first value by default
                }),
                PropertyPaneLabel('LabelSeparator', {
                  text: ' '
                }),
                PropertyPaneCheckbox('HideWebPart', {
                  text: strings.HideWebPart
                })
              ]
            },
            {
              groupName: strings.FieldsGroupLabel,
              groupFields: [
                PropertyPaneTextField('PictureUrl', {
                  label: strings.PictureUrl
                }),
                PropertyPaneTextField('LineOne', {
                  label: strings.LineOne
                }),
                PropertyPaneTextField('LineTwo', {
                  label: strings.LineTwo
                }),
                PropertyPaneTextField('LineThree', {
                  label: strings.LineThree,
                  description: strings.LineThreeDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
