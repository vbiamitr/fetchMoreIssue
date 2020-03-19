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
    private visualHOst: IVisualHost;
    private eventService: any;
    private element: HTMLElement;
    private visual: HTMLElement;

    constructor(options: VisualConstructorOptions) {
        console.log('Visual constructor', options);
        this.element = options.element;
        this.visual = document.createElement("div");
        this.visual.classList.add("custom-visual");
        this.eventService = new EventEmitter();
        this.visualHOst = options.host;
        this.target = options.element;
        this.fetchMoreCounter = 0;
        this.updateCount = 0;
        if (document) {
            const new_p: HTMLElement = document.createElement("p");
            new_p.appendChild(document.createTextNode("Update count:"));
            const new_em: HTMLElement = document.createElement("em");
            this.textNode = document.createTextNode(this.updateCount.toString());
            new_em.appendChild(this.textNode);
            new_p.appendChild(new_em);
            this.target.appendChild(new_p);

            const new_p2: HTMLElement = document.createElement("p");
            new_p2.appendChild(document.createTextNode("Row count:"));
            const new_em2: HTMLElement = document.createElement("em");
            this.textNodeRowCount = document.createTextNode("0");
            new_em2.appendChild(this.textNodeRowCount);
            new_p2.appendChild(new_em2);
            this.target.appendChild(new_p2);
        }
        this.logger();
    }

    public update(options: VisualUpdateOptions) {
        console.log(`update Called ${this.fetchMoreCounter}`);
        const consoleMessages = [];
        consoleMessages.push(`update Called ${this.fetchMoreCounter}`);

        const { dataViews } = options;
            
            this.settings = Visual.parseSettings(options && dataViews && dataViews[0]);
            if (options.operationKind == 0 /* Create */) {
                this.fetchMoreCounter = 1;
                this.eventService.emit("clear-log");
            
            } else if (options.operationKind == 1) {
                this.fetchMoreCounter++;
            }
            if ( dataViews &&  dataViews[0] &&  dataViews[0].metadata.segment) {
                this.visualHOst.fetchMoreData();
                console.log("fetchmore Called", this.fetchMoreCounter);
                consoleMessages.push(`fetchmore Called ${this.fetchMoreCounter}`);
            } 

            consoleMessages.push(`viewMode ${options.viewMode}`);
            consoleMessages.push(`editMode ${options.editMode}`);
            consoleMessages.push(`isInFocus ${options.isInFocus}`);
            consoleMessages.push(`operationKind ${options.operationKind}`);
            consoleMessages.push(`type ${options.type}`);
            
            console.log('Visual update', options);
            if (this.textNode) {
                this.textNode.textContent = (this.fetchMoreCounter).toString();
            }
            if (this.textNodeRowCount) {
                this.textNodeRowCount.textContent = (dataViews[0].matrix.rows.root.children.length).toString();
            }       
            this.eventService.emit("log", consoleMessages);
            debugger;
        // }
    }

    public logger() {
        let requestLogger = document.createElement("div");
    requestLogger.classList.add("request-logger");

    let requestLoggerHeader = document.createElement("span");
    requestLoggerHeader.classList.add("header");
    requestLoggerHeader.appendChild(
      document.createTextNode("Console Logger :")
    );

    let requestLoggerBody = document.createElement("span");
    requestLoggerBody.classList.add("content");
    this.eventService.addListener("clear-log", () => {
      while (requestLoggerBody.firstChild)
        requestLoggerBody.firstChild.remove();
    });
    this.eventService.addListener("log", data => {
        var div = document.createElement('div');
        div.classList.add('single-update');
        data.forEach(message => {
            div.append((message));
            div.append(document.createElement("br"));
        });
      requestLoggerBody.append(div);
    });

    requestLogger.appendChild(requestLoggerHeader);
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
}