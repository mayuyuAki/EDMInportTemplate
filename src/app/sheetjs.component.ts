/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
import { Component } from '@angular/core';

import * as XLSX from 'xlsx';

type AOA = any[][];
type BOB = any[][];

@Component({
	selector: 'sheetjs',
	template: `
		<div>
			上传FSM文件<input type="file" (change)="onFileChange(true,$event)" multiple="false" />	
		</div>
		<div>
			上传EDM文件<input type="file" (change)="onFileChange(false,$event)"  multiple="false" />
		</div>
		<button (click)="export()">生成EDM导入模板</button>
	`
})

export class SheetJSComponent {
	FSMData: AOA = [ [1, 2], [3, 4] ];
	EDMData: AOA = [ [0, 2], [0, 4] ];
	wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
	fileName: string = 'EDM.xlsx';

	onFileChange(edmOrFSM:any,evt: any) {
		/* wire up file reader */
		if(edmOrFSM){
			const target: DataTransfer = <DataTransfer>(evt.target);
			if (target.files.length !== 1) throw new Error('Cannot use multiple files');
			const reader: FileReader = new FileReader();
			reader.onload = (e: any) => {
				/* read workbook */
				const bstr: string = e.target.result;
				const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});

				/* grab first sheet */
				const wsname: string = wb.SheetNames[0];
				const ws: XLSX.WorkSheet = wb.Sheets[wsname];

				/* save data */
				this.FSMData = <AOA>(XLSX.utils.sheet_to_json(ws, {header: 1}));
			};
			reader.readAsBinaryString(target.files[0]);
		}
		else{
			const EDMTarget: DataTransfer = <DataTransfer>(evt.target);
			if (EDMTarget.files.length !== 1) throw new Error('Cannot use multiple files');
			const reader: FileReader = new FileReader();
			reader.onload = (e: any) => {
				/* read workbook */
				const edmbstr: string = e.target.result;
				const edmwb: XLSX.WorkBook = XLSX.read(edmbstr, {type: 'binary'});

				/* grab first sheet */
				const edmwsname: string = edmwb.SheetNames[0];
				const edmws: XLSX.WorkSheet = edmwb.Sheets[edmwsname];

				/* save data */
				this.EDMData = <BOB>(XLSX.utils.sheet_to_json(edmws, {header: 1}));
			};
			reader.readAsBinaryString(EDMTarget.files[0]);
		}
		console.log(this.FSMData);
		console.log(this.EDMData);

	}

	export(): void {
		/* generate worksheet */
        for(let i =0;i<this.FSMData.length;i++){
			for(let j =0;j<this.EDMData.length;j++){
				if(this.FSMData[i][6] === this.EDMData[j][0]){
					if(this.FSMData[i][14] && (this.FSMData[i][14] !='')){
						this.EDMData[j][47] = this.FSMData[i][14];
					}else{
						this.EDMData[j][47] = this.FSMData[i][13];
					}
				}
			}
		}
		const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.EDMData);
		/* generate workbook and add the worksheet */
		const wb: XLSX.WorkBook = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

		/* save to file */
		XLSX.writeFile(wb, this.fileName);
	}
}
