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
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import DataView = powerbi.DataView;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import { EventEmitter } from "events";

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

    constructor(options: VisualConstructorOptions) {
        console.log('Visual constructor', options);
        this.element = options.element;
        this.visual = document.createElement("div");
        this.visual.classList.add("custom-visual");
        this.eventService = new EventEmitter();
        this.visualHost = options.host;
        this.target = options.element;
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
        this.logger();
    }

    public update(options: VisualUpdateOptions) {
        const consoleMessages = [];
        consoleMessages.push(`update Called: ${this.fetchMoreCounter}`);

        const { dataViews } = options;
            
        this.settings = Visual.parseSettings(options && dataViews && dataViews[0]);
        if (options.operationKind == powerbi.VisualDataChangeOperationKind.Create) {
            this.fetchMoreCounter = 1;        
        } else if (options.operationKind == powerbi.VisualDataChangeOperationKind.Append) {
            this.fetchMoreCounter++;
        }
        
        if ( dataViews &&  dataViews[0] &&  dataViews[0].metadata.segment) {
            const fetchMoreRes = this.visualHost.fetchMoreData();
            consoleMessages.push(`fetchmore Called: ${this.fetchMoreCounter}`);
            consoleMessages.push(`fetchmore status: ${fetchMoreRes}`);
        } else {
            consoleMessages.push(`fetchmore Called: --`);
            consoleMessages.push(`fetchmore status: --`);
        }

        consoleMessages.push(`viewMode: ${options.viewMode}`);
        consoleMessages.push(`editMode: ${options.editMode}`);
        consoleMessages.push(`isInFocus: ${options.isInFocus}`);
        consoleMessages.push(`operationKind: ${options.operationKind}`);
        consoleMessages.push(`type: ${options.type}`);
        consoleMessages.push(`timestamp: ${(new Date()).toISOString()}`);
        
        if (this.textNode) {
            this.textNode.textContent = (this.fetchMoreCounter).toString();
        }
        if (this.textNodeRowCount) {
            const measures = dataViews[0].matrix.valueSources.length;
            const rowCount = this.getMemberCount(dataViews[0].matrix.rows.root);
            const columnCount = this.getMemberCount(dataViews[0].matrix.columns.root)/measures;
            this.textNodeRowCount.textContent = ` Row Count: ${rowCount.toLocaleString()} :: Column Count: ${columnCount.toLocaleString()} :: Measures: ${measures.toLocaleString()}`;
        }       
        this.addLog(consoleMessages);
    }

    public clearLogHandler = () => {
        const logContent = document.querySelector(".content");
        if(logContent) {
            while (logContent.firstChild) {
                logContent.firstChild.remove();
            }
        }
    }

    public addLog(data: string []) {
        const logContent = document.querySelector(".content");
        var div = document.createElement('div');
        div.classList.add('single-update');
        const docFragment = new DocumentFragment();
        data.forEach((message: string) => {
            const logNode = document.createElement("span");
            logNode.textContent = message; 
            docFragment.appendChild(logNode);
            docFragment.appendChild(document.createElement("br"));
        });
        div.appendChild(docFragment);
        logContent.append(div);
    }

    public logger() {
        let requestLogger = document.createElement("fieldset");
        requestLogger.classList.add("request-logger");

        let loggerTitle = document.createElement("legend");
        loggerTitle.classList.add("header");
        loggerTitle.innerText = "Console Logger";
        requestLogger.appendChild(loggerTitle);

        let clearLogBtn = document.createElement("button");
        clearLogBtn.classList.add("clear-log");
        clearLogBtn.innerText = "Clear Logs";
        loggerTitle.appendChild(clearLogBtn);
        clearLogBtn.addEventListener("click", this.clearLogHandler);

        let requestLoggerBody = document.createElement("div");
        requestLoggerBody.classList.add("content");
        requestLogger.appendChild(requestLoggerBody);

        this.visual.appendChild(requestLogger);
        this.element.append(this.visual);
    }

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