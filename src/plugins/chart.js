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
			range: [3, 24],
			bgColor: bgColors.start,
			borderType: (col) =>
				col === 24 ? borderStyle.right_side : borderStyle.top_bottom,
		},
		// { range: [70, 93], bgColor: bgColors.a2, borderType: null },
		{
			range: [25, 48],
			bgColor: bgColors.a2,
			// borderType: (col) => (col === 25 ? borderStyle.left_side : null),
		},
		{ range: [49, 54], bgColor: bgColors.SRT, borderType: null },
		{ range: [55, 60], bgColor: bgColors.SRT2, borderType: null },
		{ range: [61, 62], bgColor: bgColors.LT, borderType: null },
		{ range: [63, 68], bgColor: bgColors.bo, borderType: null },
		{ range: [69, 74], bgColor: bgColors.aided, borderType: null },
		{ range: [75, 86], bgColor: bgColors.a2, borderType: null },
		{ range: [87, 88], bgColor: bgColors.pta_avr, borderType: null },
		// { range: [114, 119], bgColor: bgColors.bo_left, borderType: null },

		// { range: [94, 99], bgColor: bgColors.SRT, borderType: null },
		// { range: [100, 105], bgColor: bgColors.SRT2, borderType: null },
		// { range: [106, 107], bgColor: bgColors.LT, borderType: null },
		// { range: [108, 113], bgColor: bgColors.bo, borderType: null },
		// { range: [120, 120], bgColor: null, borderType: null },
		// { range: [121, 122], bgColor: bgColors.pta_avr, borderType: null },
		// { range: [114, 119], bgColor: bgColors.bo_left, borderType: null },
		// {
		// 	range: [57, 64],
		// 	bgColor: bgColors.IADL,
		// 	borderType: (col) =>
		// 		col === 57
		// 			? borderStyle.left_side
		// 			: col === 64
		// 			? borderStyle.right_side
		// 			: borderStyle.top_bottom,
		// },
		// {
		// 	range: [65, 69],
		// 	bgColor: bgColors.GDS,
		// 	borderType: (col) =>
		// 		col === 69 ? borderStyle.right_side : borderStyle.top_bottom,
		// },
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
