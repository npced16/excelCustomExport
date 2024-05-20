import {
	backgroundColors as bgColors,
	textColor,
	defaultCellStyle,
	borderStyle,
} from "@/plugins/style";

/**
 * Header 셀 설정
 * @param {*} workSheet  workSheet
 * @param {*} cellRef cellr 번호 ex) A1,QO2
 * @param {*} value  표시할 내용
 * @param {*} options  bold , border , alignRight , merge
 */
export function setupTitleCell(workSheet, cellRef, value, options = {}) {
	const cell = workSheet.getCell(cellRef);
	cell.value = value;
	if (options.bold) {
		toBoldText(cell);
	}
	if (options.border) {
		setCellBorder(cell, options?.border);
	}
	if (options.alignRight) {
		cell.alignment = { horizontal: "right" };
	}
	if (options.alingleft) {
		cell.alignment = { horizontal: "left" };
	}
	if (options.merge) {
		workSheet.mergeCells(cellRef + ":" + options.merge);
	}
}
export function numberToColumnText(n) {
	let column = "";
	while (n > 0) {
		let remainder = (n - 1) % 26;
		column = String.fromCharCode(65 + remainder) + column;
		n = Math.floor((n - 1) / 26);
	}
	return column + "1";
}

// text bold 하게 만드는 함수 (기존 폰트 유지하며)
export function toBoldText(cell) {
	cell.font = { ...cell.font, bold: true };
}
// 지정된 색으로 text 칠하는 함수
export function setCellFill(cell, color) {
	if (color) {
		cell.fill = {
			type: "pattern",
			pattern: "solid",
			fgColor: { argb: color },
		};
	}
}
// 주어진 style 로 cell의 border 설정
export function setCellBorder(cell, style) {
	if (style) {
		cell.border = { ...style.border };
	}
}
//  change text color
export function setCellColor(cell, color) {
	cell.font = { ...cell.font, color: { argb: color } };
}

// title 첫줄 의
export function determineReservationsTitleStyles(colNumber) {
	// title 연결하고 boerder 설정 하는 곳
	const styleConditions = [
		{
			range: [1, 2],
			bgColor: bgColors.basic,
			borderType: borderStyle.right_side,
		},
		{
			range: [3, 21],
			bgColor: bgColors.start,
			borderType: (col) =>
				col === 21 ? borderStyle.right_side : borderStyle.top_bottom,
		},

		{
			range: [22, 65],
			bgColor: "DBDBDB",
			borderType: (col) =>
				col === 65 ? borderStyle.right_side : borderStyle.top_bottom,
		}, // Y1:BO1

		{ range: [66, 73], bgColor: bgColors.IADL, borderType: null }, // BP1:BW1
		{ range: [74, 75], bgColor: bgColors.GDS, borderType: null }, // BX1:BY1
		{ range: [76, 76], bgColor: bgColors.GDS, borderType: null }, // BZ1
		{ range: [77, 100], bgColor: bgColors.GDS_KR, borderType: null }, // CA1:DV1
		{ range: [101, 106], bgColor: bgColors.SRT, borderType: null }, // DW1:EB1
		{ range: [107, 112], bgColor: bgColors.SRT2, borderType: null }, // EC1:EH1
		{ range: [113, 114], bgColor: bgColors.IADL, borderType: null }, // EI1:EJ1
		{ range: [115, 120], bgColor: bgColors.aided, borderType: null }, // EK1:EP1
		{ range: [121, 126], bgColor: bgColors.aided_left, borderType: null }, // EQ1:EV1
		{ range: [127, 138], bgColor: bgColors.SRT2, borderType: null }, // EW1:FL1
		{ range: [139, 140], bgColor: bgColors.pta_avr, borderType: null }, // FM1:FN1
	];
	for (const condition of styleConditions) {
		const [start, end] = condition.range;
		if (colNumber >= start && colNumber <= end) {
			return {
				bgColor:
					typeof condition.bgColor === "function"
						? condition.bgColor(colNumber)
						: condition.bgColor,
				borderType:
					typeof condition.borderType === "function"
						? condition.borderType(colNumber)
						: condition.borderType,
			};
		}
	}
	const defaultStyle = {
		bgColor: bgColors.default,
		borderType: borderStyle.bottom,
	};
	return defaultStyle;
}
