import powerbi from "powerbi-visuals-api";

export interface IChunkState {
    [level: number]: number;
}

export interface IChunkConfig {
    operationKind: powerbi.VisualDataChangeOperationKind;
    done: boolean;
}

interface IParams {
    root: powerbi.DataViewMatrixNode,
    measures: powerbi.DataViewMetadataColumn [],
    level: number;
    chunkState: IChunkState;
    parents: string [];
}

export class DataParser {
    private parsedDS: any;
    public originalDS: any
    private chunkState: IChunkState;
    public tableRows: string [][];
    public rowsProcessed: number;
    constructor() {

    }

    setChunkState(dataView: powerbi.DataView, chunkConfig: IChunkConfig) {
        if(chunkConfig.done) {
            this.chunkState = null;
            return;
        }

        const rows = dataView.matrix.rows;
        const levels = rows.levels.length;
        this.chunkState = {};
        let root = rows.root;
        for(let i=0; i<levels; i++) {
            const lastIndex = root.children.length - 1;
            this.chunkState[i] = lastIndex;
            root = root.children[lastIndex];
        }
    }

    parseChunk(dataView: powerbi.DataView[], chunkConfig: IChunkConfig): [number, number] {
        this.originalDS = dataView;
        const rows = dataView[0].matrix.rows;
        let prevChunkState = chunkConfig.operationKind === powerbi.VisualDataChangeOperationKind.Append ? this.chunkState : null;        
        this.setChunkState(dataView[0], chunkConfig); 
        
        if(!prevChunkState) {
            prevChunkState = rows.levels.reduce((chunkState, levelInfo, i) => {
                chunkState[i] = 0;
                return chunkState;
            }, {});
            this.tableRows = [];
        }
        
        this.rowsProcessed = 0;
        const measures = dataView[0].matrix.valueSources;
        const params = {
            root: rows.root,
            measures,
            level: 0,
            chunkState: { ...prevChunkState },
            parents: []
        };

        const rowStart = this.tableRows.length;
        this.processRows(params);
        const rowEnd = this.tableRows.length; 
        console.log("Rows Processed = ", this.rowsProcessed);       
        return [rowStart, rowEnd];        
    }

    processRows({ root, measures, level, chunkState, parents }: IParams) {
        if(root.children) {
            const sliceIndex = chunkState[level] && chunkState[level+1] == undefined ? chunkState[level]+1 : chunkState[level];
            chunkState[level] = 0;
            root.children.slice(sliceIndex).forEach((childNode, i) => {
                const pList = parents.slice();
                pList.push(<string>childNode.levelValues[0].value);
                this.processRows({
                    root: childNode,
                    measures,
                    level: level + 1,
                    chunkState,
                    parents: pList
                });
            });
        } else {
            const values = root.values;
            const periods = Object.keys(values).length / measures.length;

            for(let i=0; i<periods; i++) {
                const row = parents.slice();
                row.push(`${i+1}`);
                measures.forEach((m, index) => {
                    row.push(<string>values[i*measures.length+index].value);
                });
                this.tableRows.push(row);
                ++this.rowsProcessed;
            }     
            
        }
    }
}