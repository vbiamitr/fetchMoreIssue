/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the ""Software""), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
*/
"use strict";

import "core-js/stable";
import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VirtualizedList from "virtualized-list";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import DataView = powerbi.DataView;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import { DataParser } from "./DataParser";
import { VisualSettings } from "./settings";

export class Visual implements IVisual {
    private target: HTMLElement;
    private updateCount: number;
    private settings: VisualSettings;
    private textNode: Text;
    private textNodeRowCount: Text;
    private fetchMoreCounter: number;
    private visualHost: IVisualHost;
    private eventService: any;
    private element: HTMLElement;
    private visual: HTMLElement;
    private dataView: any;
    private dataParser: DataParser;
    private dataTable: HTMLDivElement;

    constructor(options: VisualConstructorOptions) {
        console.log('Visual constructor', options);
        this.element = options.element;
        this.visual = document.createElement("div");
        this.visual.classList.add("custom-visual");
        this.visualHost = options.host;
        this.target = options.element;
        this.dataParser = new DataParser();
        this.fetchMoreCounter = 0;
        this.updateCount = 0;
        if (document) {
            const new_p: HTMLElement = document.createElement("p");
            new_p.appendChild(document.createTextNode("Update Count: "));
            const new_em: HTMLElement = document.createElement("span");
            this.textNode = document.createTextNode(this.updateCount.toString());
            new_em.appendChild(this.textNode);
            new_p.appendChild(new_em);
            this.target.appendChild(new_p);

            const new_p2: HTMLElement = document.createElement("p");
            new_p2.appendChild(document.createTextNode("Details:"));
            const new_em2: HTMLElement = document.createElement("b");
            this.textNodeRowCount = document.createTextNode("");
            new_em2.appendChild(this.textNodeRowCount);
            new_p2.appendChild(new_em2);
            this.target.appendChild(new_p2);
        }

        this.initTable();
        //this.logger();
    }

    public initTable() {
        if(!this.dataTable) {
            // const table = document.createElement("table");
            // table.classList.add("data-tbl");
            // table.setAttribute("border", "1");
            // this.dataTable = table;

            const tblWrapper = document.createElement("div");
            tblWrapper.classList.add("tbl-wrapper");
            this.dataTable = tblWrapper;
            //tblWrapper.appendChild(table);
            this.target.appendChild(tblWrapper);

        } else {
            while(this.dataTable.firstChild) {
                this.dataTable.firstChild.remove();
            }
        }        
    }

    public update(options: VisualUpdateOptions) {
        const { dataViews } = options;

        if(![powerbi.VisualUpdateType.All, powerbi.VisualUpdateType.Data].includes(options.type)) {
            this.updateTable();
            return;
        }

        if (options.operationKind == powerbi.VisualDataChangeOperationKind.Create) {
            this.fetchMoreCounter = 1; 
            this.initTable();       
        } else if (options.operationKind == powerbi.VisualDataChangeOperationKind.Append) {
            this.fetchMoreCounter++;
        }

        let chunkProcessed: [number, number] = [null, null];
        if (dataViews &&  dataViews[0] &&  dataViews[0].metadata.segment) {
            this.visualHost.fetchMoreData();
            chunkProcessed = this.dataParser.parseChunk(dataViews, { operationKind: options.operationKind, done: false });              
        } else {
            chunkProcessed = this.dataParser.parseChunk(dataViews, { operationKind: options.operationKind, done: true });
        }

        if (this.textNode) {
            this.textNode.textContent = (this.fetchMoreCounter).toString();
        }
        if (this.textNodeRowCount) {
            const measures = dataViews[0].matrix.valueSources.length;
            const rowCount = this.getMemberCount(dataViews[0].matrix.rows.root);
            const columnCount = this.getMemberCount(dataViews[0].matrix.columns.root)/measures;
            this.textNodeRowCount.textContent = `Total Rows Loaded: ${(rowCount * columnCount).toLocaleString()} :: Row Count: ${rowCount.toLocaleString()} :: Column Count: ${columnCount.toLocaleString()} :: Measures: ${measures.toLocaleString()} :: ChunkProcessed: [${chunkProcessed.join()}]`;
        }  

        this.updateTable();
    }

    public updateTable() {
        const tableRows = this.dataParser.tableRows;
        this.initTable();
        const vizList = new VirtualizedList(this.dataTable, {
            height: this.dataTable.getBoundingClientRect().height,
            rowCount: tableRows.length,
            rowHeight: 40,
            renderRow: (index: number) => {
                const row = tableRows[index];
                const contents = [`${index + 1}`].concat(row);
                const tbl = document.createElement("div");
                tbl.classList.add("data-tbl");
                contents.forEach((content) => {
                    const td = document.createElement("div");
                    td.classList.add("cell");
                    td.textContent = content;
                    tbl.appendChild(td);
                });
                return tbl;
            }
        });
    }

    // public ___update(options: VisualUpdateOptions) {
    //     const consoleMessages = [];
    //     consoleMessages.push(`update Called: ${this.fetchMoreCounter}`);

    //     const { dataViews } = options;
            
    //     this.settings = Visual.parseSettings(options && dataViews && dataViews[0]);
    //     if (options.operationKind == powerbi.VisualDataChangeOperationKind.Create) {
    //         this.fetchMoreCounter = 1;        
    //     } else if (options.operationKind == powerbi.VisualDataChangeOperationKind.Append) {
    //         this.fetchMoreCounter++;
    //     }
        
    //     if ( dataViews &&  dataViews[0] &&  dataViews[0].metadata.segment) {
    //         this.dataParser.parseChunk(dataViews, { operationKind: options.operationKind, done: false });  
    //         const fetchMoreRes = this.visualHost.fetchMoreData();
    //         // consoleMessages.push(`fetchmore Called: ${this.fetchMoreCounter}`);
    //         // consoleMessages.push(`fetchmore status: ${fetchMoreRes}`);
    //     } else {
    //         this.dataParser.parseChunk(dataViews, { operationKind: options.operationKind, done: true });
    //         // consoleMessages.push(`fetchmore Called: --`);
    //         // consoleMessages.push(`fetchmore status: --`);
    //     }

    //     consoleMessages.push(`viewMode: ${options.viewMode}`);
    //     consoleMessages.push(`editMode: ${options.editMode}`);
    //     consoleMessages.push(`isInFocus: ${options.isInFocus}`);
    //     consoleMessages.push(`operationKind: ${options.operationKind}`);
    //     consoleMessages.push(`type: ${options.type}`);
    //     consoleMessages.push(`timestamp: ${(new Date()).toISOString()}`);
        
    //     if (this.textNode) {
    //         this.textNode.textContent = (this.fetchMoreCounter).toString();
    //     }
    //     if (this.textNodeRowCount) {
    //         const measures = dataViews[0].matrix.valueSources.length;
    //         const rowCount = this.getMemberCount(dataViews[0].matrix.rows.root);
    //         const columnCount = this.getMemberCount(dataViews[0].matrix.columns.root)/measures;
    //         this.textNodeRowCount.textContent = `Total Rows Loaded: ${(rowCount * columnCount).toLocaleString()} :: Row Count: ${rowCount.toLocaleString()} :: Column Count: ${columnCount.toLocaleString()} :: Measures: ${measures.toLocaleString()}`;
    //     }       
    //     this.addLog(consoleMessages);
    // }

    // public clearLogHandler = () => {
    //     const logContent = document.querySelector(".content");
    //     if(logContent) {
    //         while (logContent.firstChild) {
    //             logContent.firstChild.remove();
    //         }
    //     }
    // }

    // public addLog(data: string []) {
    //     const logContent = document.querySelector(".content");
    //     var div = document.createElement('div');
    //     div.classList.add('single-update');
    //     data.forEach((message: string) => {
    //         const logNode = document.createElement("span");
    //         logNode.textContent = message; 
    //         div.appendChild(logNode);
    //         div.appendChild(document.createElement("br"));
    //     });
    //     logContent.append(div);
    // }

    // public logger() {
    //     let requestLogger = document.createElement("fieldset");
    //     requestLogger.classList.add("request-logger");

    //     let loggerTitle = document.createElement("legend");
    //     loggerTitle.classList.add("header");
    //     loggerTitle.innerText = "Console Logger";
    //     requestLogger.appendChild(loggerTitle);

    //     let clearLogBtn = document.createElement("button");
    //     clearLogBtn.classList.add("clear-log");
    //     clearLogBtn.innerText = "Clear Logs";
    //     loggerTitle.appendChild(clearLogBtn);
    //     clearLogBtn.addEventListener("click", this.clearLogHandler);

    //     let requestLoggerBody = document.createElement("div");
    //     requestLoggerBody.classList.add("content");
    //     requestLogger.appendChild(requestLoggerBody);

    //     this.visual.appendChild(requestLogger);
    //     this.element.append(this.visual);
    // }

    private static parseSettings(dataView: DataView): VisualSettings {
        return <VisualSettings>VisualSettings.parse(dataView);
    }

    /**
     * This function gets called for each of the objects defined in the capabilities files and allows you to select which of the
     * objects and properties you want to expose to the users in the property pane.
     *
     */
    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
        return VisualSettings.enumerateObjectInstances(this.settings || VisualSettings.getDefault(), options);
    }

    private getMemberCount(root: powerbi.DataViewMatrixNode): number {
        let memberCount = 0;
        const parseDimension = (root: powerbi.DataViewMatrixNode) => {
            if(root.children) {
                root.children.forEach(parseDimension);
            } else {
                memberCount+=1;
            }
        }
        parseDimension(root);
        return memberCount;
    }
 }