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
			上传FSM文件<input type="file" (change)="onFileChange($event)" multiple="false" />
		</div>
		<button (click)="export()">生成EDM导入模板</button>
	`
})

export class SheetJSComponent {
	FSMData: AOA = [ [1, 2], [3, 4] ];
	EDMData: BOB = [ [0, 2], [0, 4] ];
	wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
	fileName: string = 'EDM.xlsx';

	onFileChange(evt: any) {
		/* wire up file reader */
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
		console.log(this.FSMData);

	}

	export(): void {
		/* generate worksheet */
		// EDM Excel COLUMN Name
		//this.EDMData[0] = ["SOLD TO", "BP Role", "客户名称", "Customer English Name", "客户简称", "Search term 2", "Market Segment", "L4 Cust Hier Node", "L4 Cust Hier Name", "Name 2", "Date", "FH Customer", "FH TM Code", "FH TM", "FH ADM", "FH DM", "Standard", "Template", "City", "Postl Code", "TranspZone", "Street", "Street 2", "Rg", "RelCat", "SHIP TO", "BP Role", "Standard", "Name 1", "Search term 1", "门店名称", "Central", "Name 2", "Last name", "First name", "City-发货地址", "Postl Code", "TranspZone", "Street 2", "Rg", "Telephone", "Street - 发货地址", "GT address", "集团代码", "集团名称", "联系人信息", "联系人电话", "TM", "客户市场类型"];
		this.EDMData[0] = ["SOLD TO","客户名称", "客户简称", "SHIP TO","门店名称","City-发货地址","Street - 发货地址","集团代码", "集团名称", "联系人信息", "联系人电话", "TM", "客户市场类型"];
		for(let i =0;i<this.FSMData.length-2;i++){
			this.EDMData.push([]);
		}
		if(this.FSMData[0][3] === 'SoldTo号'){
			//Sold To
			for(let i =1;i<this.FSMData.length;i++){
				this.EDMData[i][0] = this.FSMData[i][3];
				this.EDMData[i][1] = this.FSMData[i][4];
				this.EDMData[i][2] = this.FSMData[i][5];
				this.EDMData[i][3] = '';
				this.EDMData[i][4] = this.FSMData[i][5];
				this.EDMData[i][5] = this.FSMData[i][7];
				this.EDMData[i][6] = this.FSMData[i][8];
				this.EDMData[i][7] = this.FSMData[i][9];
				this.EDMData[i][8] = this.FSMData[i][10];
				this.EDMData[i][9] = this.FSMData[i][11];
				this.EDMData[i][10] = this.FSMData[i][12];
				if(this.FSMData[i][14] && (this.FSMData[i][14] !='')){
					this.EDMData[i][11] = this.FSMData[i][14];
				}else{
					this.EDMData[i][11] = this.FSMData[i][13];
				}
				this.EDMData[i][12] = this.FSMData[i][15];
			}
			console.log(this.EDMData.length);

		}else{
			//Ship To
			for(let i =1;i<this.FSMData.length;i++){
				this.EDMData[i][0] = '';
				this.EDMData[i][1] = this.FSMData[i][4];
				this.EDMData[i][2] = this.FSMData[i][5];
				this.EDMData[i][3] = this.FSMData[i][3];
				this.EDMData[i][4] = this.FSMData[i][5];
				this.EDMData[i][5] = this.FSMData[i][7];
				this.EDMData[i][6] = this.FSMData[i][8];
				this.EDMData[i][7] = this.FSMData[i][9];
				this.EDMData[i][8] = this.FSMData[i][10];
				this.EDMData[i][9] = this.FSMData[i][11];
				this.EDMData[i][10] = this.FSMData[i][12];
				if(this.FSMData[i][14] && (this.FSMData[i][14] !='')){
					this.EDMData[i][11] = this.FSMData[i][14];
				}else{
					this.EDMData[i][11] = this.FSMData[i][13];
				}
				this.EDMData[i][12] = this.FSMData[i][15];
			}
			console.log(this.EDMData.length);

		}
		const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.EDMData);
		/* generate workbook and add the worksheet */
		const wb: XLSX.WorkBook = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

		/* save to file */
		XLSX.writeFile(wb, this.fileName);
	}
}
