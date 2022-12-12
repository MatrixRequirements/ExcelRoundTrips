
import{ PluginCore,IPluginConfig} from "./core/PluginCore"
import { IProjectSettings } from "./Interfaces";
/** This class is allows you to configure the features of your plugin.
 * 
 *  You can also implement functions to into the plugin (at start in the constructor, when loading a project, when loading an item)
 * 
     */
export class Plugin extends PluginCore {

    
    /**This part enables which 
     * 
     * See IPluginConfig interface for explanation of parameters
    */
    
    static config: IPluginConfig = {
        /*  Page in admin client to configure settings across all projects - set enabled to false if not needed. 
            The page itself is implemented in the _ServerSettingsPage.ts 
        */
        customerSettingsPage: {
            id: "ERTCustomerSettings",
            title: "ERT customer settings page",
            type: "ERTcs",
            enabled: false,
            defaultSettings: {
                myServerSetting: "default value for setting defined in Interfaces.ts",
            },
            settingName: "ERT_settings",
            help: "This is my help text",
            helpUrl:"https://docs23.matrixreq.com"
        },
        /*  Page in admin client to configure settings for one specific project - set enabled to false if not needed.
            The page itself is implemented in the _ProjectSetingsPage.ts 
        */
        projectSettingsPage: {
            id: "ERTprojectsettings",
            title: "ERT projectsettings page",
            type:"ERTps",
            enabled: true,
            defaultSettings: {
                rules:[]
            },
            settingName: "ERT_settings",
            help: "This is my help text",
            helpUrl:"https://docs23.matrixreq.com"
        },
        /*  Add an entry in the tool menu of an item or folder - set enabled to false if not needed.
            The tool itself is implemented in the _Tool.ts 
        */
        menuToolItem: {
            enabled: true,
            title:"Sync From Excel",
        },
        /*  Add a custom field to enter some data in the UI - set enabled to false if not needed.
            The field itself is implemented in the _Control.ts 
        */
        field: {
            enabled: false,
            fieldType: "excelroundtrips",
            title: "excelroundtrips-field",
            fieldConfigOptions: {
                id: "excelroundtrips",
                capabilities: {
                    canBePublished: false,
                    canBeReadonly: true,
                    canBeXtcPreset: false,
                    canHideInDoc: false,
                    canBeUsedInDocs: false,
                    canRequireContent: true,
                },
                class: "",
                help: "",
                label: "excelroundtrips-field",
            }
        },
        /*  Add a dashboard inside a project - set enabled to false if not needed.
            The field itself is implemented in the _Control.ts 
        */
        dashboard: {        
            
            id:"ERT",
            title: "ERT dashboard page",
            enabled: false,
            icon: "fal fa-cog",
            parent: "DASHBOARDS",
            usefilter: true,
            order: 9999,
        }
    };

    /**
     * The constructor is loaded once after all the source code is loaded by the browser. 
     * Nothing is known about the instance, project, user etc.
     * You can use it to initialize things like callbacks after item changes
     */
    constructor() {
        super();
        
        // here is a good place to register callbacks for UI events (like displaying or saving items)
    }

    /**
     * This method is called each time  a project has been loaded and initialized. 
     * At the time it is called, all project settings, categories etc are defined.
     * 
     * @param project // id of the loaded project
     */
    onInitProject(project:string) {

        // here is a good place to decide based on the project (maybe some project setting), whether the plugin should be enabled 
        
        // if not:
        // this.enabledInContext = false;
    }

    /** this method is called just before the rendering of an item is done
    * It is also called when opening the create item dialog. 
    * 
    * 
    * @param _item: the item which is being loaded in the UI 
    */
    onInitItem(item: IItem) {
        
        // here is a good place to decide based on the selection in the tree, whether the plugin should be enabled 
        
        let projectConfig = <IProjectSettings>IC.getSettingJSON( Plugin.config.projectSettingsPage.settingName, {});
        if ( !projectConfig.rules || projectConfig.rules.filter( rule => rule.category == item.type ).length == 0) {
            this.enabledInContext = false;
        } else {
            this.enabledInContext = true;
        } 
    }
}

declare global {
    
    // eslint-disable-next-line no-unused-vars
    interface Window {
        plugins: unknown;
    }
}
    $(() => {
    // Register the plugin
    $(function () {
        window.plugins.register(new Plugin());
    });
});