// eslint-disable-next-line no-unused-vars
import { Plugin } from "./Main";
import { IProjectSettingMapping, IProjectSettings } from "./Interfaces"; 


interface IImportColumn {
    label:string, // as displayed
    id:string, // id in json
    isLabel?:boolean, // true if column is a label
    
    index?:number,
    fieldId?:number,
    fieldType?:string,
}
interface IImportRow {
    cells:string[]
}

interface IStringNeedleMap {
    [key: string]: XRTrimNeedleItem;
}
export class Tool {
    private duplicates:string[] = [];
    public isDefault = true;
    private xml:JQuery;
    private dlg:JQuery;
   
    private jsonFields = ["richtext",  "user" , "date", "text", "textline" ,"test_result",
        "crosslinks", "gateControl","fileManager", "reviewControl"];

    private newImportRoot:string;
    private projectSettingMapping:IProjectSettingMapping;

    /** callback to show or hide the menu for a selected item or folder
    * 
    * */ 
    showMenu(itemId:string) {
        return itemId.startsWith("F-");
    }

    /** callback when user executes the custom the menu entry added to items or folders 
     * 
     * */ 
    menuClicked(itemId:string) {
        /* Insert code here */
        this.newImportRoot = itemId;
        this.startSync();
    }

  
    // start import wizard
    protected startSync() {
        let that= this;

        // get rules for that category (we assume there's exactly one matching rule)

        let projectConfig = <IProjectSettings>IC.getSettingJSON( Plugin.config.projectSettingsPage.settingName, {});
        that.projectSettingMapping = projectConfig.rules.filter( rule => rule.category == ml.Item.parseRef(that.newImportRoot).type )[0];

        // show dialog
        that.dlg = $("<div>").appendTo($("body"));
        let ui = $("<div style='height:100%;width:100%'>");
        
        ml.UI.showDialog( that.dlg, "Sync Excel rows with items", ui, $(document).width() * 0.9, app.itemForm.height() * 0.9,
        [{
            text: 'Next',
            class: 'btnDoIt',
            click:  () => { /* nothing to do */}
            }], 1, true, true, () => { that.dlg.remove(); }, () => { that.wizardStepPrepare(ui, $(".btnDoIt", that.dlg.parent())); }, () => { }
        );    
    }

    /* *************************************************
        Wizard
    ************************************************* */

    // import excel one WS at a time mapping column to fields
    private wizardStepPrepare( ui:JQuery, next:JQuery ) {
        
        ui.html("<h1>Step 1: Prepare the excel</h1>");
        let ol = $("<ol>").appendTo(ui);
        $("<li>Unmerge all cells in all worksheets</li>").appendTo(ol);
        $("<li>Remove all comments</li>").appendTo(ol);
        
        ui.append("<h1>Step 2: Upload the file</h1>");
    
        this.appendFileUpload(ui, next,  (ui, next) => this.wizardStepSelectWS(ui, next));

        ml.UI.setEnabled(next,false);
    }

    // add a section to upload an excel, which is converted to xml and handed to the next step
    private appendFileUpload(ui:JQuery, next:JQuery, nextStep:(ui:JQuery,next:JQuery) => void) {
        let that = this;

        $("<div>").appendTo(ui).fileManager( {
            parameter:{
                "readonly":false, "manualOnly":true, "extensions":[
                    ".xls",
                    ".xlsx"],
                "single":true,
                "textTodo":" ",
                hideNoFileInfo:true
            },
            controlState: ControlState.FormEdit,
            canEdit: true,
            help: " ",
            fieldValue: "[]",
            valueChanged: function() { /* nothing to do */ },
            processExternally: function(files:FileList) {
                ml.File.convertXLSXAsync( files[0]).done(function(text:string) {
                
                    // "fix" not exported cells <Cell />
                    text  = text.replace(/<Cell \/>/g,"<Cell></Cell>");
                    that.xml = $(text);
                    if (that.xml.length == 0) {
                        ml.UI.showError("Cannot read xml", "Conversion failed somehow...");
                        return;
                    }
                    let worksheets = $("Worksheet", that.xml);
                    if (worksheets.length==0)  {
                        ml.UI.showError("No worksheets in xml", "There are no worksheets in xml...");
                        return;
                    }
                    nextStep(ui,next);
                });
                return false;
            }
        });
    }

    // for normal file excel to item:
    private wizardStepSelectWS( ui:JQuery, next:JQuery ) {
        let that = this;
        ui.html("");
        ui.closest(".ui-dialog-content").removeClass("dlg-v-scroll");
        let h1 = $("<h1>Step 3: Select the worksheet to import</h1>").appendTo(ui);
        let thead:JQuery;
        let tbody:JQuery;
        let select = $("<div>").appendTo(ui);
        let tableContainer = $("<div>").appendTo(ui);
        let worksheets = $("Worksheet", this.xml);
        let ws:IDropdownOption[] = [];
        $.each(worksheets, function( idx, worksheet){
            ws.push({id:""+idx, label:$(worksheet).attr("ss:Name")});
        });
        
        select.mxDropdown({
            controlState: ControlState.FormEdit,
            canEdit: true,
            help: "",
            fieldValue: "",
            valueChanged: function () {
                if (select) {
                    let selected = select.getController().getValue();
                    tableContainer.html("");
                    let table = $("<table class='table table-bordered'>").appendTo(tableContainer);
                    thead = $("<thead>").appendTo(table);
                    thead.append("<th style='padding:0'>include</th>");
                    tbody = $("<tbody>").appendTo(table);
                    
                    ml.UI.setEnabled(next,selected!="");
                    if (selected) {
                        that.rows=[];
                        
                        let wsRows = $("row", worksheets[selected]);

                        // count max columns and copy data
                        let maxColumns = 0;
                        $.each(wsRows, function(rowIdx, row) {
                            let cells:string[] = [];
                                
                            let cellIdx = 0;
                            $(row).children("cell").each( (cidx, cell) => {
                                cellIdx = that.getCellIndex($(cell), cellIdx);
                                cells[cellIdx] = $(cell).text().trim();
                                maxColumns = Math.max( maxColumns, cellIdx);
                                cellIdx++;
                            });
                            that.rows.push( {cells:cells});
                            
                        });
                        // create header
                        maxColumns++;
                        for( let cellIdx=0;cellIdx<maxColumns; cellIdx++) {
                            thead.append("<th style='padding:0'>");
                        }
                        // create table content
                        let uid = that.getColumnIndexFromExcelColumn( that.projectSettingMapping.uidColumn );
                        
                        $.each(that.rows, function(rowIdx, row){
                            let checkedInclude = ( (uid==-1 || row.cells[uid]) &&  (!that.projectSettingMapping.excludeUpTo || rowIdx>=that.projectSettingMapping.excludeUpTo))?"checked":"";
                            
                            let tr = $("<tr>").appendTo(tbody);
                            $(`<td><input type='checkbox' ${checkedInclude}></td>` ).appendTo(tr);
                            for( let cellIdx=0; cellIdx<maxColumns; cellIdx++) {
                                tr.append("<td>"+(row.cells[cellIdx]?row.cells[cellIdx]:"")+"</td>");
                            }
                        });
                    } 
                }
            },
            parameter: {
                placeholder: 'select worksheet',
                maxItems: 1,
                options: ws,
                groups: [],
                create: false,
                sort: false,
                splitHuman:false
            }
        });
       
      
        ml.UI.setEnabled(next,false);
        next.unbind( "click" );
        next.click(  function() {
            select.remove();
            that.wizardStepMapColumns(ui,next, h1, thead, tbody);
        });
    }

  
    private getCellIndex(cell:JQuery, order:number):number {
        // it's either the place (or an explicit index if there are some gaps...)
        if ($(cell).attr("ss:Index")) {
            return Number($(cell).attr("ss:Index"))-1;
        }
        return order;
    }
    private wizardStepMapColumns( ui:JQuery, next:JQuery, h1:JQuery, thead:JQuery, tbody:JQuery) {
        let that = this;
        h1.html("Step 4: Map columns to fields");
        
        // get options for mapping (supported fields)
        let ddOptions:IImportColumn[]= [ {isLabel:false,label:"ignore",id:""}, {isLabel:false,label:"HIERARCHY", id:"HIERARCHY"}, {isLabel:false,label:"FOLDER",id:"FOLDER"}, {isLabel:false,label:"TITLE",id:"TITLE"} ];

        let category =  ml.Item.parseRef(app.getCurrentItemId() ).type;
        $.each(IC.getFields(category), function(idx, field) {
            if (field.fieldType == "checkbox" || field.fieldType == "dropdown" || that.jsonFields.indexOf( field.fieldType) !=-1 
                || field.fieldType == "steplist"|| field.fieldType == "test_steps" || field.fieldType == "test_steps_result" )  {
                // simple direct copy fields
                ddOptions.push( {isLabel:false,label:field.label,id:field.label});
            } else  if (field.fieldType == "risk2" ) {
                // risk field
                let rc = (field.parameterJson && (<IRiskParameter>field.parameterJson).riskConfig)?((<IRiskParameter>field.parameterJson).riskConfig):IC.getRiskConfig();
                $.each( rc.factors, function( factorIdx, factor) {
                    let opt = field.label + "." + factor.type;
                    ddOptions.push( {isLabel:false,label:opt,id:opt});
                    $.each( factor.weights, function( weightIdx,weight) {
                        let opt = field.label + "." + weight.type;
                        ddOptions.push( {isLabel:false,label:opt,id:opt});
                    });
                });
                if (rc.postReduction && rc.postReduction.weights) {
                    $.each(  rc.postReduction.weights, function( weightIdx,weight) {
                        let opt = field.label + ".#." + weight.type;
                        ddOptions.push( {isLabel:false,label:opt,id:opt});
                    });
                }
            }
        });

        $.each( ml.LabelTools.getLabelDefinitions([category]), function( labelIdx, label) {
            ddOptions.push({isLabel:true, label:"label: " +label.label+ " ("+  ml.LabelTools.getDisplayName(label.label) + ")", id:label.label});
        });

        // map the excel columns to fields to indices with fields
        let tableColumn:INumberStringMap = {};
        if (that.projectSettingMapping.columnToFieldMap) {
            for (  let excelColumn of Object.keys(that.projectSettingMapping.columnToFieldMap )) {
                let indexColumn = that.getColumnIndexFromExcelColumn( excelColumn );
                if (indexColumn==-1) {
                    console.log(`Column definition is not correct "${excelColumn}" needs to be a column in excel style like  "A" or "B" ... `)
                } else {
                    tableColumn[indexColumn] = that.projectSettingMapping.columnToFieldMap[excelColumn];
                }
            }
        }

        // map the label columns to fields to indices with fields
        let labelColumn:INumberStringMap = {};
        if (that.projectSettingMapping.columnToLabelMap) {
            for (  let excelColumn of Object.keys(that.projectSettingMapping.columnToLabelMap )) {
                let indexColumn = that.getColumnIndexFromExcelColumn( excelColumn );
                if (indexColumn==-1) {
                    console.log(`Column definition is not correct "${excelColumn}" needs to be a column in excel style like  "A" or "B" ... `)
                } else {
                    labelColumn[indexColumn] = that.projectSettingMapping.columnToLabelMap[excelColumn];
                }
            }
        }
        // map the title column
        let titleColum = 1 + that.getColumnIndexFromExcelColumn( that.projectSettingMapping.titleColumn );
        
        // add the drop down and select the defaults
        that.columns = [];
        $.each( $("th", thead), function(thIdx, th) {
            if (thIdx>0) {
                let dd = $("<select style='width:100%'>").appendTo(th).data("index", thIdx);
                $.each(ddOptions, function(ddOptionIdx, ddOption) {
                    let selected =  (ddOption.id == tableColumn[thIdx-1])?"selected": // column is mapped field
                                    (ddOption.id == labelColumn[thIdx-1])?"selected": // column is mapped label
                                    (ddOption.id == "TITLE" && thIdx == titleColum)?"selected":"";
                    
                    $(`<option data-index='${thIdx-1}' data-islabel='${(ddOption.isLabel?1:0)}' ${selected}>${ddOption.label}</option>`).appendTo(dd).val(ddOption.id);
                });    
            }
        });

        // function to map UI choices to column mapping
        let evalSelection = () => {
            that.columns = [];
            $.each( $( "select option:selected", thead ), function( selIdx, selected) {
                if ( $(selected).text() ) {
                    that.columns.push({isLabel: $(selected).data("islabel")=="1", id:<string>$(selected).val(), label:$(selected).text(), index:$(selected).data("index") });
                }
            });
            ml.UI.setEnabled(next,that.columns.length !=0);
        };
        // react on UI changes
        $("select", thead).on("change", ( event:JQueryEventObject) => {
            evalSelection();
            return null;
        });
        ml.UI.setEnabled(next,false);
        // do it once after drawing
        evalSelection();

       next.unbind( "click" );
        next.click(  function() {
            that.rows = that.rows.filter( function( row, rowIdx) { return $($("tr input", tbody)[rowIdx]).is(":checked")});
            that.import(  ui, next );
        });
    }

    /* *************************************************
        MassImport
    ************************************************* */
    
    private columns:IImportColumn[];
    private folderColumn:number;
    private hierarchyColumn:number;
    private titleColumn:number;

    private rows:IImportRow[];
    private messages: JQuery;
    private hierarchyMap:IStringMap;

    protected import(  ui:JQuery, next:JQuery ) {
        let that = this;
        
        ui.html("<h1>Step 5: Retrieving current items</h1>");
        ml.UI.getSpinningWait("getting data").appendTo(ui);
        
        ml.UI.setEnabled(next, false);

        let category = ml.Item.parseRef(that.newImportRoot).type;

        let uidMap:IStringNeedleMap = {};
        let uidFieldName = that.projectSettingMapping.columnToFieldMap[that.projectSettingMapping.uidColumn?that.projectSettingMapping.uidColumn:"zzzzzz"];
        let uidFieldId = IC.getFieldByName( category, uidFieldName).id;
                
        restConnection.getProject(`needle?search=mrql:category=${category}&labels=1&fieldsOut=*`).done( (results:XRGetProject_Needle_TrimNeedle) => { 
            for (let needle of results.needles) {
                let uidFields  = needle.fieldVal.filter( fv => fv.id == uidFieldId);
                if (uidFields.length) {
                    uidMap[uidFields[0].value] = needle; 
                }
            }
            that.importUpdate( ui, next, uidMap, uidFieldName );
        });

    }

    // import one worksheet in excel -> each row one item
    private importUpdate(  ui:JQuery, next:JQuery, uidMap:IStringNeedleMap, uidFieldName:string ) {
        let that = this;
        
        ui.html("<h1>Step 6: Converting ....</h1>");
        this.messages = $("<ul>").appendTo(ui);
        ml.UI.setEnabled(next, false);

        let category = ml.Item.parseRef(this.newImportRoot).type;
        let fields = IC.getFields(category);
        if (fields.length==0) { 
            this.error(-1, "Category does not exist/has no fields");
            return;
        }
        // map columns to fields
        this.folderColumn = -1;
        this.titleColumn = -1;
        this.hierarchyColumn = -1;

        // ignore columns which do not need to be matched
        this.columns = this.columns.filter( function (column) { return column.label != "ignore";});

        // add matching fields id's and types
        $.each( this.columns, function( idx:number, column:IImportColumn ) {
            if ( column.id == "HIERARCHY") {
                that.hierarchyColumn = column.index;
            } else if ( column.id == "FOLDER") {
                that.folderColumn = column.index;
            } else if ( column.id == "TITLE") {
                that.titleColumn = column.index;
            } else if (column.isLabel) {
            } else {
                let field = fields.filter( function(field) { return field.label.toLowerCase() == column.id.toLowerCase().split(".")[0];});
                if (field.length != 1) { ml.Logger.log( "Info", "the field " + column.id.toLowerCase().split(".")[0] + " does not exist in category " + category);return;}
                column.fieldId = field[0].id;
                column.fieldType = field[0].fieldType;
            }
        });

        if (this.folderColumn !=-1 && this.hierarchyColumn!=-1) {
            this.error(-1, "Error: folder options can only be hierarchical or flat - not both.");
            return;
        }
        // get rid of special columns
        this.columns = this.columns.filter( function (column) { return column.id != "FOLDER" && column.id != "HIERARCHY" && column.id != "TITLE";});

        // go through all rows...
        this.importUpdateData( category, this.newImportRoot, uidMap, uidFieldName).done( function() {
            ml.UI.showSuccess( "All Done!");
            that.messages.append("<li><b>Done!</b></li>");
            ml.UI.setEnabled(next, true);
        }).fail( function() {
            that.messages.append("<li><b>Failed!</b></li>");
            ml.UI.showError( "Conversion Failed", "see log");
            ml.UI.setEnabled(next, true);
        });
        next.html("Close");
        next.unbind( "click" );
        next.click(  function() {
            $(".ui-dialog-titlebar-close", that.dlg.parent()).trigger("click");
        });
    }
    // import or update one row from excel -> each row one item
    private importUpdateData ( category:string, root:string, uidMap:IStringNeedleMap, uidFieldName:string) {
        let that = this;
        let res = $.Deferred();
        
        if (this.hierarchyColumn == -1) {
            return this.importUpdateAllRows( category, root, root, 0, uidMap, uidFieldName);
        } else {
            let neededFolders:string[] = [];
            $.each( this.rows, function(rowIdx, row) {
                let hier = row.cells[that.hierarchyColumn];
                // hier can be A | B | C | D: a folder D in a folder C in a folder B...
                // make a flat list of needed folders A, A | B, A | B | C, ... in case they done exist yet...
                if ( hier ) {
                    let parts = hier.split("|");
                    let addPath = "";
                    $.each( parts, function(partIdx, part) {
                        addPath += part;
                        if ( neededFolders.indexOf(addPath)==-1) {
                            neededFolders.push(addPath);
                        }
                        addPath += "|";
                    });
                }
            });
            that.hierarchyMap = {};
            that.createFolderHierarchy(category, root, neededFolders, 0).done(function () {
                that.importUpdateAllRows( category, root, root, 0, uidMap, uidFieldName).done(function () {
                    res.resolve();
                }).fail( function() {
                    res.reject();
                });
            }).fail( function() {
                res.reject();
            });
        }
        return res;
    }
    
    private createFolderHierarchy ( category:string, root:string, neededFolders:string[], rowIdx:number) {
        let that = this;
        let res = $.Deferred();

        if (rowIdx >= neededFolders.length ) {
            res.resolve();
            return res;
        }
        // convert A | B | C to parent = A | B and last = C
        let full = neededFolders[rowIdx].split ("|");
        let last = full.splice(full.length-1,1) [0];
        let parent = full.join("|");
        // parent should already exist (ensured by called) if not it's in the root
        let parentId = parent?that.hierarchyMap[parent]:root;
        let item:IItemPut = {
            title:last.trim(),
            children:[]
        }
        app.createItemOfTypeAsync( category, item, "import", parentId).done( function(newFolder) {
            // create the lookup
            that.hierarchyMap[neededFolders[rowIdx]] = newFolder.item.id;

            that.showCreatedItem( rowIdx, newFolder.item.id, newFolder.item.title);

            // ml.UI.showSuccess( "created folder: " + last);
            that.createFolderHierarchy ( category, root, neededFolders, rowIdx+1).done( function() {
                res.resolve();
            }).fail( function() {
                res.reject();
            });
        }).fail( function() {
            res.reject();
            that.error(rowIdx, "creating folder failed - aborting");
        });

        return res;
    }

    private importUpdateAllRows( category:string, root:string, current:string, rowIdx:number, uidMap:IStringNeedleMap, uidFieldName:string) {
        let that = this;
        let res = $.Deferred();

        if (rowIdx>= this.rows.length ) {
            res.resolve();
            return res;
        }

        let row = this.rows[rowIdx];

        // **********************
        // handle folder rows
        // **********************
        if ( this.folderColumn != -1 && row.cells[this.folderColumn] ) {
            ml.Logger.log( "Info", "create folder " +row.cells[this.folderColumn]  + " in " + root );
      
            let folderToCreate:IItemPut = {
                title:row.cells[this.folderColumn],
                children:[]
            }

            app.createItemOfTypeAsync( category, folderToCreate, "import", root).done( function(newFolder) {
                that.showCreatedItem( rowIdx, newFolder.item.id, newFolder.item.title);
                //ml.UI.showSuccess( rowIdx + "/" + (that.rows.length-1) + ": created new folder " + newFolder.item.id );
                that.importUpdateAllRows ( category, root, newFolder.item.id, rowIdx+1, uidMap, uidFieldName).done( function() {
                    res.resolve();
                }).fail( function() {
                    res.reject();
                });
            }).fail( function() {
                res.reject();
                that.error(rowIdx, "creating folder - aborting");
            });
           
            return res;
        }

        // **********************
        // handle empty rows
        // **********************
        let content = 0;
        $.each( this.columns, function( colIdx, column ) {
            if ( row.cells[column.index]) content++;  
        });  
        if ( that.titleColumn!=-1 &&  row.cells[that.titleColumn])  content++;
        if (!content) {
            that.warning(rowIdx, "skipping row " + (rowIdx+1));
            that.importUpdateAllRows ( category, root, current, rowIdx+1, uidMap, uidFieldName).done( function() {
                res.resolve();
            }).fail( function() {
                res.reject();
            });
            return res;
        }

        // **********************
        // import update items
        // **********************

        let uidIndex = that.getColumnIndexFromExcelColumn( that.projectSettingMapping.uidColumn );

        // fill the item

        let item:IItemPut = {
            title:(this.titleColumn!=-1 && row.cells[this.titleColumn])?this.rows[rowIdx].cells[this.titleColumn]:("ROW " + (rowIdx +1))
        }
        let labels:string[] = [];
        // create a fake dummy UI - only used for logic in risks' right now
        let riskField = 0;
        let riskControlName = "";
        let dummy = $("<div>");
        let newItem =new ItemControl({
            control: dummy,
            controlState: ControlState.DialogCreate,
            parent: current,
            type: category,
            isItem: true,
            changed: function () { /* empty on purpose */}});
        let postWeights:IRiskValueFactorWeight[];
        
        $.each( this.columns, function( colIdx, column ) {
            let cell = row.cells[column.index]?row.cells[column.index]:"";

            if (column.fieldType=="risk2") {
                riskField = column.fieldId;
                riskControlName = column.label.split(".")[0];
                if ( column.label.split(".")[1] == "#" ) {
                    // post weights
                    if (!postWeights) postWeights = [];
                    postWeights.push( {description:"", label:"", type:column.label.split(".")[2], value: Number(cell)});
                } else {
                    let input =  column.label.split(".")[1];
                    that.setRiskInput( dummy, input, cell, false, rowIdx);
                    (<RiskControlImpl>newItem.getControlByName( riskControlName ).getController()).riskChange();
                }
            } else  if (column.fieldType=="richtext") {
                (<IStringMap>item)[column.fieldId] = cell.replace(/\n/g, "<br>");
            } else  if (  that.jsonFields.indexOf( column.fieldType) !=-1 ) { 
                (<IStringMap>item)[column.fieldId] = cell;
            } else  if (column.fieldType=="checkbox") {
                (<IBooleanMap>item)[column.fieldId] =cell.toLowerCase() == "x" ||  cell.toLowerCase() == "true" || cell == "1";
            } else  if (column.fieldType=="dropdown") {
                that.makeDropdown(category, column, rowIdx, cell, item);
            } else  if (column.fieldType=="test_steps" || column.fieldType=="test_steps_result" || column.fieldType=="steplist") {
                that.makeTable(category, column, rowIdx, cell, item);
            } else if (column.isLabel)  {
                if (cell.toLowerCase() == "x" ||  cell.toLowerCase() == "true" || cell == "1") {
                    labels.push( column.id );
                }
            } else {
                that.warning(rowIdx, "unsupported field type: " + column.fieldType + " cannot be converted automatically");
            }
        });

        if (that.projectSettingMapping.dirtyLabel) {
            labels.push(that.projectSettingMapping.dirtyLabel);
        }
        if (labels.length) {
            item.labels = labels.join(",");
        }
        if (riskField) {
            let riskVal:IRiskValue = <IRiskValue>JSON.parse(newItem.getControlByName(riskControlName).getController().getValue());

            if (postWeights) riskVal.postWeights = postWeights;
            (<IStringMap>item)[riskField] = JSON.stringify(riskVal);
        }

        let folder = current;
        if (that.hierarchyColumn!=-1 &&  row.cells[that.hierarchyColumn] && that.hierarchyMap[row.cells[that.hierarchyColumn]]) {
            folder = that.hierarchyMap[row.cells[that.hierarchyColumn]];
        }

        // either update or create the item
            
        let uid = (uidIndex >=0 && row.cells[uidIndex])?row.cells[uidIndex]:"";
        if (uid && uidMap[uid]) {
            // this is actually an update


            // fix risk controls and mitigations
            if (riskField) {
                let newRisk = JSON.parse(item[riskField]);
                let oldRisks = uidMap[uid].fieldVal.filter( fv => fv.id==riskField );
                if (oldRisks.length) {
                    // the risk existed before
                    let oldRisk = JSON.parse( oldRisks[0].value );
                    newRisk.mitigations = oldRisk.mitigations; // these cannot be imported anyway
                    if (newRisk.postWeights && oldRisk.postWeights) {
                        for (let newPW of newRisk.postWeights ) {
                            let oldPWs = oldRisk.postWeights.filter( pw => pw.type == newPW.type && pw.value == newPW.value);
                            if (oldPWs.length==1) {
                                newPW.description = oldPWs[0].description;
                            }
                        }
                    }
                }
                item[riskField] = JSON.stringify(newRisk);

            }

            that.updateIfNeeded( rowIdx, uidMap[uid], category, item).done( () => {
                that.importUpdateAllRows ( category, root, current, rowIdx+1, uidMap, uidFieldName).done( function() {
                    res.resolve();
                }).fail( function() {
                    res.reject();
                });
            });
        } else {
            // create a new item

            app.createItemOfTypeAsync( category, item, "import", folder?folder:current).done( function(newItem) {

                that.showCreatedItem( rowIdx, newItem.item.id, newItem.item.title);
                that.importUpdateAllRows ( category, root, current, rowIdx+1, uidMap, uidFieldName).done( function() {
                    res.resolve();
                }).fail( function() {
                    res.reject();
                });
            }).fail( function() {
                that.error(rowIdx, "creating item");
                res.reject();
            });
        }
        return res;
    }

    private updateIfNeeded(rowIdx:number, serverVersion:XRTrimNeedleItem, category:string, newVersion: IItemPut) {
        let that = this;
        let res = $.Deferred();

        // prepare update
        let update:IItemPut = {
            title:newVersion.title,
            id: ml.Item.parseRef( serverVersion.itemOrFolderRef ).id,
            onlyThoseFields:1,
            onlyThoseLabels:1
        };

        let changed = false;

        // check title
        if (serverVersion.title != newVersion.title) {
            changed = true;
        } 

        // check labels
        let labelChanges:string[] = [];
        let labelsServer = serverVersion.labels?serverVersion.labels.split(","):[];
        let labelsNew = newVersion.labels?newVersion.labels.split(","):[];
        // for all mapped labels
        for (let label of that.columns.filter( c => c.isLabel ).map(l => l.id)) {
            let isServer = labelsServer.indexOf( "("+label+")")!=-1;
            let isNew = labelsNew.indexOf( label )!=-1;
            if (isServer && !isNew) {
                changed = true;
                labelChanges.push( "-"+label ); // remove the label
            } else if (!isServer && isNew) {
                changed = true;
                labelChanges.push( label ); // add the label
            } 
        }
       
        // check fields
        let fields = Object.keys( newVersion ).filter( key => key != "title" && key != "labels");
        for( let field of fields) {
            let exists = serverVersion.fieldVal.filter( fv => fv.id +"" == field && fv.value == newVersion[field]).length != 0;
            if (!exists && newVersion[field]) {
                changed = true;
                update["fx"+field] = newVersion[field];
            }
        }
        
        that.showUpdatedItem( rowIdx, serverVersion.itemOrFolderRef, changed?"CHANGED":"SAME");
        if (changed) {

            if (that.projectSettingMapping.dirtyLabel) {
                // mark as dirty
                labelChanges.push(that.projectSettingMapping.dirtyLabel);
            }
            if (labelChanges.length) {
                update.labels = labelChanges.join(",");
            }
            app.updateItemInDBAsync(update, "import" ).done( function (updatedItem) {
                res.resolve();
            }).fail( () => {
                that.error(rowIdx, "updating item " + update.id);
            })
        } else {
            res.resolve();
        }
        
        return res;
    }
    private makeTable(category: string, column: IImportColumn, rowIdx: number, cell: string, item: IItemPut) {
        let fieldConfig  = IC.getFields(category).filter(function (field) { return field.id == column.fieldId; })[0];

        let columnNames:string[] = [];
        if (fieldConfig.fieldType == "test_steps" || fieldConfig.fieldType == "test_steps_result") {
            $.each( IC.getTestConfig().render[category].columns, function(cidx,columnDef) {
                columnNames[cidx+1] = columnDef.field;
            });
        } else if ( fieldConfig.fieldType == "steplist") {
            $.each(fieldConfig.parameterJson.columns, function(cidx, columnDef) {
                columnNames[(<number>cidx)+1] = columnDef.field;
            });
        } 

        let data:IStringMap[] = [];

        $.each( cell.split("|#"), function(ridx, row) {
            let jsonRow:IGenericMap = {};

            $.each( row.split("|*"), function(cidx, column) {
                if (columnNames[cidx]) {
                    let text = column.trim();
                    jsonRow[columnNames[cidx]]=text;
                }
            });

            data.push(jsonRow);
        });

        (<IStringMap>item)[column.fieldId] = JSON.stringify(data);
    }
    private makeDropdown(category: string, column: IImportColumn, rowIdx: number, cell: string, item: IItemPut) {
        let that = this;

        let jsonParams = <IDropdownParams>IC.getFields(category).filter(function (field) { return field.id == column.fieldId; })[0].parameterJson;
        let dd_options = jsonParams.optionSetting;
        if (!dd_options) {
            this.warning(rowIdx, "unsupported dropdown configuration: only dropdowns with values in settings are supported!");
        }
        else {
            let dd = IC.getDropDowns(dd_options);
            if (dd.length != 1) {
                this.warning(rowIdx, "drop down config is missing: only dropdowns with values in settings are supported!");
            }
            else {
                let options = (jsonParams.maxItems>1)?cell.split("|"):[cell];
                let mapped:string[] = [];
                $.each( options, function( optIdx, option) {
                    let optionId = "";
                    $.each(dd[0].value.options, function (optIdx, opt) {
                        if (opt.label.toLowerCase() == option.trim().toLowerCase()) {
                            optionId = opt.id;
                        }
                    });
                    if (!optionId) {
                        if (jsonParams.create) {
                            optionId =  option.trim();
                        } else {
                            that.warning(rowIdx, "drop down option does not exist: '" + option.trim() + "'");
                        }
                    } 
                    if (  mapped.length < jsonParams.maxItems ||  mapped.length == 0) {
                        mapped.push(optionId);
                    } else {
                        that.warning(rowIdx, "ignored option: '" + optionId + "' - to many selected options for this config!");
                    }
                });
                
                (<IStringMap>item)[column.fieldId] = mapped.length?mapped.join():"";
            }
        }
    }

    setRiskInput( item:JQuery, name:string, value:string, post:boolean, row:number ) {
        if ( $("input[name="+name+"]" , item).length > 0) {
            $("input[name="+name+"]", item).val( value );
        } else if ( $("div[name="+name+"]", item).length  > 0 ) {
            $("div[name="+name+"]", item).data( "realValue", value );
        } else if ( $("textarea[name="+name+"]", item).length  > 0 ) {
            $("textarea[name="+name+"]", item).val( value );
        } else if ( $("select[name="+name+"]", item).length  > 0 ) {
            // dropdown
            let select =$(post?$("select[name="+name+"]", item)[1]:$("select[name="+name+"]", item)[0]);
            let optionGoood = false;
            $.each( $("option", select), function( idx, option) {
                if ( $(option).data("value")==value ) {
                    $(option).prop("selected", true); 
                    optionGoood = true;
                } else {
                    $(option).prop("selected", false);
                }
            });
            // second change, go by the attr value
            if (!optionGoood && value) { 
                $.each( $("option", select), function( idx, option) {
                    if ( $(option).attr("value")==value ) {
                        $(option).prop("selected", true); 
                        optionGoood = true;
                    } else {
                        $(option).prop("selected", false);
                    }
                });
            }
            // third change, go by the text
            if (!optionGoood && value) { 
                $.each( $("option", select), function( idx, option) {
                    if ( $(option).text()==value ) {
                        $(option).prop("selected", true); 
                        optionGoood = true;
                    } else {
                        $(option).prop("selected", false);
                    }
                });
            }
            if (!optionGoood) { 
                this.warning( row,  "risk input "+ name +" is a select, but the option '" + value + "' does not exist");
            }
        } else { 
            this.warning( row,  "risk input "+ name +" does not exist as input, text area or select");
        }
    }

    private error( row:number, msg:string ) {
        this.messages.append( "<li style='color:red'>Error: (row " + (row+1) + ") " + msg + "</li>");
        return msg;
    }

    private warning( row:number, msg:string) {
        
        if ( this.duplicates.indexOf(msg) == -1 ) {
            this.duplicates.push(msg);

            this.messages.append( "<li style='color:red'>Warning: (row " + (row+1) + ") " + msg + "</li>");
            ml.UI.showError("Warning", "row " + (row+1) + ": " + msg);
        }
        return msg;
    }

    private showCreatedItem(  row:number, id:string, title:string) {
        if (row==-1) {
            this.messages.append( "<li>Created " + id + " " + title + "</li>");
        } else {
            this.messages.append( "<li>Created (row " + (row+1) + ") " + id + " " + title + "</li>");
        }
    }
    private showUpdatedItem(  row:number, id:string, title:string) {
        this.messages.append( "<li>Update (row " + (row+1) + ") " + id + " " + title + "</li>");
    }


    /** Excel: A,B,C -> index 0,1,2 */
    protected getColumnIndexFromExcelColumn( column:string):number {
        let columnPos = 0;
        for(let p = 0; p < column.length; p++){
            columnPos = column.charCodeAt(p) - 64 + columnPos * 26;
        }
        return columnPos-1;
    }

   
}
 
