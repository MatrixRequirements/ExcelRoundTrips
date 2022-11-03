import { IPluginSettingPage } from "./core/PluginCore";
import { IProjectSettings } from "./Interfaces";
import { Plugin } from "./Main";

    /* project Setting page closure*/
    export function ProjectSettingsPage():IPluginSettingPage <IProjectSettings>{
        let self: IPluginSettingPage<IProjectSettings> = {};

        if (window["ConfigPage"] !== undefined) {
            self = { ...Object.getPrototypeOf(new ConfigPage()) }
        }

        
        self.getSettingsDOM = (settings:IProjectSettings): JQuery => {
            
            return $(`
                <div class="panel-body-v-scroll fillHeight">
                    <div>sorry no UI yet - the syntax is as follows:</div> 
                    <pre>
export interface IProjectSettings {
    /** a list of import/sync rules */
    rules:IProjectSettingMapping[]; 
}

export interface IProjectSettingMapping {
    /**  each category can have a different rule, so this specifies the category for the rule below */
    category:string; 
    /** excel column name which has the unique ID (A,B, ...).  */
    uidColumn:string;
    /** excel column name which has the title (A,B, ...).  */
    titleColumn:string;
    /** how many rows to exclude from excel (normally should be at least to exclude a header row) */
    excludeUpTo:Number;
    /** set this label if item has changed */
    dirtyLabel:string;
    /** maps a column to a field (or property of risk ), e.g. {"A":"legacy id", "B":"Description", "AE":"Risk.harm"} */
    columnToFieldMap:IStringMap 
    /** maps a column to a label id */
    columnToLabelMap:IStringMap 
}
                    </pre>
                    <div id="controls"></div>
                </div>
                `);
        };


        self.settings = () => {
            let currentSettings = {};
            if (window["configApp"] !=undefined) {
                let filterSettings = configApp.getJSONProjectSettings(self.getProject(), Plugin.config.projectSettingsPage.settingName);
                if (filterSettings.length == 1)
                    currentSettings = filterSettings[0].value;
            }
            else {
                currentSettings = IC.getSettingJSON(Plugin.config.projectSettingsPage.settingName, {});
            }
            console.log("Returning project settings");
            return { ...Plugin.config.projectSettingsPage.defaultSettings, ...currentSettings }
        };
        self.renderSettingPage = () => {
            self.initPage(
                `${ Plugin.config.projectSettingsPage.title}` ,
                true,
                undefined,
                Plugin.config.projectSettingsPage.help,
                Plugin.config.projectSettingsPage.helpUrl,
                undefined
            );
            self.showSimple();
        };
        self.saveAsync = ()=> {
            let def =  configApp.setProjectSettingAsync(self.getProject(), Plugin.config.projectSettingsPage.settingName, JSON.stringify(self.settingsChanged), configApp.getCurrentItemId());
            def.done(() => {
                self.settingsOriginal = { ...self.settingsChanged };
                self.renderSettingPage();
            })
            return def;
        }
        self.getProject = () => {
            /* get the project id from the setting page */
            return configApp.getCurrentItemId().split("-")[0];
        }
        self.showAdvanced = () => {
            console.debug("Show advanced clicked");
            self.showAdvancedCode(JSON.stringify(self.settingsChanged), function (newCode: string) {
                self.settingsChanged = JSON.parse(newCode);
                self.paramChanged();
               
            });
        };
        self.showSimple = () => {

            self.settingsOriginal = self.settings();
            self.settingsChanged = { ...self.settingsOriginal };
            let dom = self.getSettingsDOM(self.settingsChanged);
            // ml.UI.addTextInput($("#controls",dom), "My Project setting", self.settingsChanged, "myProjectSetting",self.paramChanged);
            app.itemForm.append(dom);
        };

        self.paramChanged = () => {
            configApp.itemChanged(JSON.stringify(self.settingsOriginal) != JSON.stringify(self.settingsChanged));
        }
        return self;
    }
