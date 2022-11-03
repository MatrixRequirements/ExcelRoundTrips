/* Setting interfaces */

/**
 * This file defines all the data structures which might be shared between UI components and printing
 * 
 */

    /** Server setting for plugin.
    * 
    * This you can use to save setting on an instance level (for all projects)
    * The user can edit these in the admin through the Server Setting Page
*/
export interface IServerSettings {
    /** Server Setting example */
    myServerSetting: string;
}      

/** Project setting for plugin
* 
* This you can use to save setting for one specific project.
* The user can edit these in the admin through the Project Setting Page
*/
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
    excludeUpTo:number;
    /** set this label if item has changed */
    dirtyLabel:string;
    /** maps a column to a field (or property of risk ), e.g. {"A":"legacy id", "B":"Description", "AE":"Risk.harm"} */
    columnToFieldMap:IStringMap 
    /** maps a column to a label id */
    columnToLabelMap:IStringMap 
}

/** Setting for custom fields 
* 
* These allow a user to add parameters to custom field defined by the plugin
* each time it is added to a category
*/
export interface IPluginFieldParameter extends IFieldParameter {
    /** see below */
    options: IPluginFieldOptions;
}

/**  interface for the configuration options of field */
export interface IPluginFieldOptions  {
    // to be defined
}

/** interface for the value to be stored by custom field */
export interface IPluginFieldValue {
    // to be defined
}

/** this allows to store parameters for printing 
* 
* This parameters can be overwritten in the layout and are used by the custom section printing
*/
    export interface IPluginPrintParams extends IPrintFieldParams {
    class:string // default:"". additional class for outermost container
}