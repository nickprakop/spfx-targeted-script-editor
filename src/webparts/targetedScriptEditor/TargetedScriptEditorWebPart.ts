
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { IPropertyPaneConfiguration, IPropertyPaneField, PropertyPaneToggle } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { Placeholder } from '@pnp/spfx-controls-react';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import * as strings from 'TargetedScriptEditorWebPartStrings';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import spservices from '../../services/spservices';

export interface ITargetedScriptEditorWebPartProps {
  description: string;
  scriptBody: string;
  spPageContextInfo: boolean;
  teamsContext: boolean;
  targetedGroups: IPropertyFieldGroupOrPerson[];
  removePadding: boolean
}

export interface ITargetedScriptEditorWebPartState {
  exucuteScript: boolean;
}

export default class TargetedScriptEditorWebPart extends BaseClientSideWebPart<ITargetedScriptEditorWebPartProps> {
  public _unqiueId;
  editorProp: any;
  peoplePicker: any;


  constructor() {
    super();
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected _onConfigure = () => {
    // Context of the web part
    this.context.propertyPane.open();
  }

  public render(): void {
    
    if (this.displayMode == DisplayMode.Read) {
      if (this.properties.removePadding) {
        let element = this.domElement.parentElement;
        // check up to 5 levels up for padding and exit once found
        for (let i = 0; i < 5 && element !== null; i++) {
          const style = window.getComputedStyle(element);
          const hasPadding = style.paddingTop !== "0px";
          if (hasPadding) {
            element.style.paddingTop = "0px";
            element.style.paddingBottom = "0px";
            element.style.marginTop = "0px";
            element.style.marginBottom = "0px";
          }
          element = element.parentElement;
        }
      }
    }

    if (this.properties.scriptBody?.length > 0) {
      ReactDom.unmountComponentAtNode(this.domElement);
      if (this.properties.targetedGroups?.length > 0) {
        let proms: any[] = [];
        const errors: string[] = [];
        const _sv = new spservices();
        this.properties.targetedGroups.map((item) => {
          proms.push(_sv.isMember(item.fullName, this.context.pageContext.legacyPageContext[`userId`], this.context.pageContext.site.absoluteUrl));
        });
        void Promise.race(
          proms.map(p => {
            return p.catch(err => {
              errors.push(err);
              if (errors.length >= proms.length) {
                this.domElement.innerHTML = "";
                throw errors;
              }
              // eslint-disable-next-line @typescript-eslint/no-empty-function
              return new Promise(() => { });
            });
          })).then(val => {
            // eslint-disable-next-line @typescript-eslint/no-floating-promises
            this.executeScript(this.domElement);
          });
      } else {
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        this.executeScript(this.domElement);
      }
    } else {
      const placeHolderElement = React.createElement(Placeholder, {
        iconName: "Edit",
        iconText: "Configure your web part",
        description: "Please configure the web part.",
        buttonLabel: "Configure",
        onConfigure: this._onConfigure,
      });
      ReactDom.render(placeHolderElement, this.domElement);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected async loadPropertyPaneResources(): Promise<void> {
    //import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
    this.editorProp = await import(
      /* webpackChunkName: 'scripteditor' */
      '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor'
    );

    // import { PropertyFieldPeoplePicker, IPropertyFieldGroupOrPerson, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
    this.peoplePicker = await import(
      /* webpackChunkName: 'scripteditor' */
      '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker'
    );
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let webPartOptions: IPropertyPaneField<any>[] = [
      this.editorProp.PropertyFieldCodeEditor('scriptBody', {
        label: 'Edit Code',
        panelTitle: 'Edit Code',
        initialValue: this.properties.scriptBody,
        onPropertyChange: this.onPropertyPaneFieldChanged,
        properties: this.properties,
        disabled: false,
        key: 'codeEditorFieldId',
        language: this.editorProp.PropertyFieldCodeEditorLanguages.JavaScript
      }),
      this.peoplePicker.PropertyFieldPeoplePicker('targetedGroups', {
        label: 'Target Audience',
        initialData: this.properties.targetedGroups,
        allowDuplicate: false,
        principalType: [this.peoplePicker.PrincipalType.SharePoint],    
        onPropertyChange: this.onPropertyPaneFieldChanged,
        context: this.context as any,
        properties: this.properties,
        onGetErrorMessage: undefined,
        deferredValidationTime: 0,
        key: 'groupsFieldId'
      }),
      PropertyPaneToggle("removePadding", {
        label: "Remove top/bottom padding of web part container",
        checked: this.properties.removePadding,
        onText: "Remove padding",
        offText: "Keep padding"
      }),
      PropertyPaneToggle("spPageContextInfo", {
        label: "Enable classic _spPageContextInfo",
        checked: this.properties.spPageContextInfo,
        onText: "Enabled",
        offText: "Disabled"
      }),
    ];

    if (this.context.sdks.microsoftTeams) {
      let config = PropertyPaneToggle("teamsContext", {
        label: "Enable teams context as _teamsContexInfo",
        checked: this.properties.teamsContext,
        onText: "Enabled",
        offText: "Disabled"
      });
      webPartOptions.push(config);
    }

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupFields: webPartOptions
            }
          ]
        }
      ]
    };
  }

  private evalScript(elem) {
    const data = (elem.text || elem.textContent || elem.innerHTML || "");
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    const scriptTag: HTMLScriptElement = document.createElement("script");

    for (let i = 0; i < elem.attributes.length; i++) {
      const attr = elem.attributes[i];
      // Copies all attributes in case of loaded script relies on the tag attributes
      if (attr.name.toLowerCase() === "onload") continue; // onload handled after loading with SPComponentLoader
      scriptTag.setAttribute(attr.name, attr.value);
    }

    // set a bogus type to avoid browser loading the script, as it's loaded with SPComponentLoader
    scriptTag.type = scriptTag.src?.length > 0 ? "pnp" : "text/javascript";
    // Ensure proper setting and adding id used in cleanup on reload
    scriptTag.setAttribute("pnpname", this._unqiueId);

    try {
      // doesn't work on ie...
      scriptTag.appendChild(document.createTextNode(data));
    } catch (e) {
      // IE has funky script nodes
      scriptTag.text = data;
    }

    headTag.insertBefore(scriptTag, headTag.firstChild);
  }

  // Finds and executes scripts in a newly added element's body.
  // Needed since innerHTML does not run scripts.
  //
  // Argument element is an element in the dom.
  private async executeScript(element: HTMLElement) {
    this.domElement.innerHTML = this.properties.scriptBody;
    // clean up added script tags in case of smart re-load        
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    let scriptTags = headTag.getElementsByTagName("script");
    for (let i = 0; i < scriptTags.length; i++) {
      const scriptTag = scriptTags[i];
      if (scriptTag.hasAttribute("pnpname") && scriptTag.attributes["pnpname"].value == this._unqiueId) {
        headTag.removeChild(scriptTag);
      }
    }

    if (this.properties.spPageContextInfo && !window["_spPageContextInfo"]) {
      window["_spPageContextInfo"] = this.context.pageContext.legacyPageContext;
    }

    if (this.properties.teamsContext && !window["_teamsContexInfo"]) {
      window["_teamsContexInfo"] = this.context.sdks.microsoftTeams?.context;
    }

    // Define global name to tack scripts on in case script to be loaded is not AMD/UMD
    (<any>window).ScriptGlobal = {};

    // main section of function
    const scriptNodes = this.domElement.getElementsByTagName("script");
    const scripts: HTMLScriptElement[] = [];

    for (let i = 0; scriptNodes[i]; i++) {
      const child: HTMLScriptElement = scriptNodes[i];
      if (!child.type || child.type.toLowerCase() === "text/javascript") {
        scripts.push(child);
      }
    }

    let oldamd = null;
    if (window["define"] && window["define"].amd) {
      oldamd = window["define"].amd;
      window["define"].amd = null;
    }

    for (let i = 0; i < scripts.length; i++) {
      try {
        let script = scripts[i];
        // Add unique param to force load on each run to overcome smart navigation in the browser as needed
        if (script.src) {
          const prefix = script.src.indexOf('?') === -1 ? '?' : '&';
          let scriptUrl = script.src + prefix + 'pnp=' + new Date().getTime();
          await SPComponentLoader.loadScript(scriptUrl, { globalExportsName: "ScriptGlobal" });
        }
      } catch (error) {
        if (console.error) {
          console.error(error);
        }
      }
    }

    if (oldamd) {
      window["define"].amd = oldamd;
    }

    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.parentNode) {
        scriptTag.parentNode.removeChild(scriptTag);
      }
      this.evalScript(scripts[i]);
    }
    // execute any onload people have added
    for (let i = 0; i < scripts.length; i++) {
      let script = scripts[i];
      // Add unique param to force load on each run to overcome smart navigation in the browser as needed
      if (script.onload) {
        script.onload[i]();
      }
    }
  }
}
