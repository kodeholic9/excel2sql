const xlsx = require('xlsx')
const fs = require('fs')

////////////////////////////////////////////////////////////////////////////
//
// 데이타 파일(Excel)내 파싱 규칙을 정의
//
////////////////////////////////////////////////////////////////////////////
const excelConfig = {
	meta: {},
	columns: [],
	callbacks: {},

	// 기타
	columnKeyHashList: []
}

// config를 초기화한다.
function initConfig (
	// metaInfo
	{ dataStart, tableName, statementPrefix }, 
	// columns
	columns, 
	// callbacks
	{ genValues }
) {	
	excelConfig.meta = {
		dataStart: dataStart ?? 1,
		tableName: tableName,
		statementPrefix: statementPrefix
	}
	excelConfig.columns = columns
	excelConfig.callbacks = {
		genValues: genValues
	}

	// columnKeyHashList를 초기화한다.
	columns.forEach((x, xIdx) => {
		if (!x.keyType) return

		let columnKey = excelConfig.columnKeyHashList.find(y => y.keyType === x.keyType) 
		if (!columnKey) {
			columnKey = {
				keyType: x.keyType,
				positions: [],
				hash: {}
			}
			excelConfig.columnKeyHashList.push(columnKey)
		}
		columnKey.positions.push(xIdx)
	})

	console.log('excelConfig: ', excelConfig)
}

// data를 초기화한다.
function initData(data) {
	return data.slice(excelConfig.format.dataStart)	
}

// 에러 로그를 출력한다.
function errorLog(s) {
	console.log('* ERROR ERROR ERROR ERROR ERROR ERROR ERROR ERROR *')
	console.log('*')
	console.log('* ' + s)
	console.log('*')
	console.log('* ERROR ERROR ERROR ERROR ERROR ERROR ERROR ERROR *')
}

// 결과 출력 파일명 날짜 포맷
function getCurrentFormattedDate() {
	const now = new Date();
    
	// 년도, 월, 일, 시간, 분을 추출
	const year = now.getFullYear();
	const month = String(now.getMonth() + 1).padStart(2, '0'); // 월은 0부터 시작하므로 +1 필요
	const day = String(now.getDate()).padStart(2, '0');
	const hours = String(now.getHours()).padStart(2, '0');
	const minutes = String(now.getMinutes()).padStart(2, '0');

	// 포맷팅하여 반환
	return `${year}${month}${day}_${hours}${minutes}`;
}

// 공란 여부 체크
function checkEmptyColumns(values, index) {
	return excelConfig.columns.some((x, xIdx) => {
		if (x.nullable) return false
		if (values[xIdx] === undefined || values[xIdx] === null || values[xIdx] === '') {
			errorLog(`${x.label} 속성이 공란입니다. Line: ${excelConfig.meta.dataStart + index}`)
			return true
		}

		return false
	})
}

// 행 전체가 공란인지 체크
function checkEmptyLine(values, num) {
	for (let i = 0; i < num; i++) {
		if (values[i]) return false
	}

	return true
}

// 중복 여부 체크
function checkDuplicatedKey(values) {
	return excelConfig.columnKeyHashList.some(x => {
		//key를 생성한다.
		let key = x.positions.reduce((acc, y) => {
			acc.push(values[y])
			return acc
		}, []).join('-')

		//key가 존재하는지 체크한다.
		if (x.hash[key]) {
			errorLog(`키값[${key}] 중복이 발생하였습니다!!`)
			return true
		}
		x.hash[key] = true

		console.log(`key: ${key}`, values)

		return false
	})
}

// 공란에 대한 보상 처리
function replaceColumn(value, dType) {
	switch (dType) {
		case "string":
			if (value === undefined || value === null || value === '') return "''"
			return `'${value}'`

		case "number":
			if (value === undefined || value === null) return "NULL"
			return value
	}

	return value
}

// data를 객체로 변환한다.
function data2Obj(values) {
	return excelConfig.columns.reduce((obj, x, xIdx) => {
		obj[x.name] = replaceColumn(values[xIdx], x.dType)
		return obj
	}, {})
}

// INSERT 구문 생성
function generateSqlInsert(data) {
	const rows = data.slice(excelConfig.meta.dataStart)	

	const results = rows.reduce((acc, x, xIdx) => {
		// 빈 행 체크
		if (checkEmptyLine(x, excelConfig.columns.length)) {
			console.log('@@ SKIP AN EMPTY LINE!!')
			return acc
		}

		// 필수항목 공란 체크
		if (checkEmptyColumns(x, xIdx + 1)) {
			process.exit(1)
		}


		// 중복된 데이타 체크
		if (checkDuplicatedKey(x)) {
			process.exit(1)
		}

		// 배열를 객체로 생성
		acc.push(data2Obj(x))

		return acc
	}, [])

	if (results.length === 0) {
		errorLog('Excel파일내 유효한 데이타가 존재하지 않습니다.')
		process.exit(1)
	}

	//console.log(results)
	const sqlValues = results.map(excelConfig.callbacks.genValues).join(',')

	// 파일명 생성 및 저장
	const outputFile = excelConfig.meta.tableName + "_" + getCurrentFormattedDate() + ".sql"
	try {
		fs.writeFileSync(outputFile, excelConfig.meta.statementPrefix + sqlValues + ";")
		console.log('---------------------------------------------------------------------')
		console.log('-')
		console.log('- 총 ' + results.length + ' 건을 저장하였습니다.');
		console.log('- 파일명: ' + outputFile);
		console.log('-')
		console.log('---------------------------------------------------------------------')
	} catch (e) {
		errorLog('파일 생성중 오류가 발생하였습니다.');
		process.exit(1)
	}
}

/////////////////////////////////////////////////////////////
//
// main
//
/////////////////////////////////////////////////////////////

// 상수 정의
const TABLE_NAME = "drprg_location_subway"
const SKIP_LINE = 1 //skip header
const INSERT_STATEMENT_PREFIX = "INSERT INTO `practice`.`drprg_location_subway` (`idx`, `is_use`, `latitude`, `longitude`, `addr`, `area`, `line`, `operator`, `station`, `reg_idx`, `reg_dt`, `mod_idx`, `mod_dt`) VALUES "


// 명령줄 인자를 파싱한다.
const [,, filePath] = process.argv
if (!filePath) {
	console.error("Usage: node excel2sql.js <Excel path>")
	process.exit(1)
}

console.log(filePath)

// 엑셀 파일 읽기
const workbook = xlsx.readFile(filePath);
const sheetName = workbook.SheetNames[0]; // 첫 번째 시트 선택
const worksheet = workbook.Sheets[sheetName];

const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 }); // 2D 배열로 변환
if (!data || data.length <= 1) {
	errorLog("Excel 형식에 오류가 있습니다.")
	process.exit(1)
}

// config 초기화
initConfig(
	// excel format & table meta info
	{
		dataStart: 1,
		tableName: TABLE_NAME,
		statementPrefix: INSERT_STATEMENT_PREFIX
	}, 
	// columns
	[
		{ label: '우선순위', name: 'idx', keyType: 'index', dType: 'number' },
		{ label: '운영주체', name: 'operator', keyType: 'subway', dType: 'string' },
		{ label: '지역', name: 'area', keyType: 'subway', dType: 'string' },
		{ label: '호선번호', name: 'line', keyType: 'subway', dType: 'string' },
		{ label: '역명', name: 'station', keyType: 'subway', dType: 'string' },
		{ label: '도로주소', name: 'addr', dType: 'string' },
		{ label: '위도', name: 'latitude', nullable: true, dType: 'number' },
		{ label: '경도', name: 'longitude', nullable: true, dType: 'number' },
	],
	// callbacks
	{
		genValues: (x, xIdx) => {
			return `\n\t( ${x.idx}, 1, ${x.latitude}, ${x.longitude}, ${x.addr}, ${x.area}, ${x.line}, ${x.operator}, ${x.station}, NULL, NULL, NULL, NULL )`
		}
	}
)

// 첫 번째 행을 열 이름으로 설정
generateSqlInsert(data)

process.exit(0)

