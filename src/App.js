import './App.css';
import {useState} from 'react';
import * as XLSX from 'xlsx';

const REQUIRED_COLUMNS = [
	'Aerial Cable 192F - 192F Distribution Cable - AER',
	'Aerial Cable 144F - 144F Distribution Cable - AER',
	'Aerial Cable 48F - 48F Distribution Cable - AER',
	'Aerial Cable 4F',
	'Aerial Cable 96F - 96F Distribution Cable - AER',
	'Aerial drop box',
	'Cabinet (Existing)',
	'Closure',
	'Cross Accessory (192F)',
	'Cross Accessory (144F)',
	'Cross Accessory (48F)',
	'Cross Accessory (96F)',
	'JEPCO ONU',
	'Splitter (2:16)',
	'Splitter (2:8)',
	'Suspension Accessory (192F)',
	'Suspension Accessory (144F)',
	'Suspension Accessory (48F)',
	'Suspension Accessory (96F)',
	'Tension Accessory (192F)',
	'Tension Accessory (48F)',
	'Tension Accessory (96F)',
];

function App() {

	const [clusterSheet, setClusterSheet] = useState(null);
	const [bom, setBom] = useState(null);

	const handleClusterSheetFile = (e) => {
		const reader = new FileReader();
		reader.onload = (e) => {
			/* Parse data */
			const ab = e.target.result;
			const workBook = XLSX.read(ab, {type: 'array', cellDates: true});
			/* Get first worksheet */
			const workSheetName = workBook.SheetNames[0];
			const ws = workBook.Sheets[workSheetName];
			/* Convert array of arrays */
			const edata = XLSX.utils.sheet_to_json(ws, {header: 1});
			/* Update state */
			//   setCols(make_cols(ws["!ref"]));
			setClusterSheet(edata);
		};
		reader.readAsArrayBuffer(e.target.files[0]);
	};

	const handleBOMFile = (e) => {
		const reader = new FileReader();
		reader.onload = (e) => {
			/* Parse data */
			const ab = e.target.result;
			const workBook = XLSX.read(ab, {type: 'array', cellDates: true});
			/* Get first worksheet */
			const workSheetName = workBook.SheetNames[0];
			const ws = workBook.Sheets[workSheetName];
			/* Convert array of arrays */
			const edata = XLSX.utils.sheet_to_json(ws, {header: 1});
			/* Update state */
			//   setCols(make_cols(ws["!ref"]));
			setBom(edata);
		};
		reader.readAsArrayBuffer(e.target.files[0]);
	};

	const handleClick = () => {
		console.log(clusterSheet);
		console.log(bom);

		const mergedArray = [...clusterSheet];

		mergedArray[0] = [
				...mergedArray[0],
			'Aerial Cable 192F - 192F Distribution Cable - AER',
			'Aerial Cable 144F - 144F Distribution Cable - AER',
			'Aerial Cable 48F - 48F Distribution Cable - AER',
			'Aerial Cable 4F',
			'Aerial Cable 96F - 96F Distribution Cable - AER',
			'Aerial drop box',
			'Cabinet (Existing)',
			'Closure',
			'Cross Accessory (192F)',
			'Cross Accessory (144F)',
			'Cross Accessory (48F)',
			'Cross Accessory (96F)',
			'JEPCO ONU',
			'Splitter (2:16)',
			'Splitter (2:8)',
			'Suspension Accessory (192F)',
			'Suspension Accessory (144F)',
			'Suspension Accessory (48F)',
			'Suspension Accessory (96F)',
			'Tension Accessory (192F)',
			'Tension Accessory (48F)',
			'Tension Accessory (96F)',
		];

		let values = {};

		let AGG_ID = ""
		for (const bomElement of bom) {
			if (bomElement[0]?.includes("Sub-Area Name")) {
				AGG_ID = bomElement[2].split("=")[1].trim();
				values[AGG_ID] = [];

			} else {
				const materialIndex = REQUIRED_COLUMNS.findIndex(column => {
					return column === bomElement[1];
				});


				if (materialIndex === -1) {
					continue;
				}

				values[AGG_ID][materialIndex] = (Math.round(bomElement[6]*100))/100;
			}
		}

		for (const [aggID, quantities] of Object.entries(values)) {
			const aggIdIndex = mergedArray.findIndex(value => value[0] === parseInt(aggID));

			if (aggIdIndex === -1) {
				continue;
			}

			mergedArray[aggIdIndex] = [...mergedArray[aggIdIndex], ...quantities];

		}

		console.log(mergedArray);

		const sheetData = [];
		mergedArray.forEach((element) => {
			if (element[0] !== "DP") {
				const clusterObject = {};
				mergedArray[0].forEach((elem, idx) => {
					clusterObject[elem] = element[idx] ?? 0
				})

				sheetData.push(clusterObject);
			}
		})

		const workSheet = XLSX.utils.json_to_sheet(sheetData);
		const workbook = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(workbook, workSheet, 'Sheet1');

		const wbout = XLSX.write(workbook, { type: 'binary', bookType: 'xlsx' });

		// Convert binary string to ArrayBuffer
		const buf = new ArrayBuffer(wbout.length);
		const view = new Uint8Array(buf);
		for (let i = 0; i < wbout.length; i++) {
			view[i] = wbout.charCodeAt(i) & 0xff;
		}

		// Create Blob object
		const blob = new Blob([buf], { type: 'application/octet-stream' });

		// Create download link
		const url = URL.createObjectURL(blob);
		const a = document.createElement('a');
		a.href = url;
		a.download = 'Modified Cluster Sheet.xlsx';
		document.body.appendChild(a);
		a.click();
		document.body.removeChild(a);
		URL.revokeObjectURL(url);

	}
	return (<div className="App">
		<div className={'App-input-container'}>
			<label htmlFor={'cluster-sheet-file'} className={'App-label'}>
				Upload the cluster sheet file here
			</label>
			<br/>
			<input type={'file'} name={'cluster-sheet-file'} className={'App-input'}
			       onChange={handleClusterSheetFile}
			       accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"

			/>
		</div>
		<div className={'App-input-container'}>
			<label htmlFor={'bom-file'} className={'App-label'}>
				Upload the bom file here
			</label>
			<br/>
			<input type={'file'} name={'bom-file'} className={'App-input'}
			       onChange={handleBOMFile}
			       accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
			/>
		</div>
		<button className={'App-button'} onClick={handleClick}>
			Merge Files
		</button>
	</div>);
}

export default App;
