export const backgroundColors = {
	start: "FFCCFF",
	basic: "9999FF",
	IADL: "F8CBAC",
	GDS: "BED7EE",
	GDS_KR: "FFF2CC",
	SRT: "8FABDB", //순음
	SRT2: "DAE3F3", //순음 2
	aided: "D0CECE", // 보청기 착용
	aided_left: " E2F0D9", // 보청기 착용 왼
	pta_avr: "BF9000",
	header: "328493",
	default: "DBDBDB",
	// aided: "E2F0D9",
};
export const textColor = {
	red: "FF0000",
	white: "FFFFFF",
	default: "000000",
};
export const defaultCellStyle = {
	font: {
		name: "맑은 고딕",
		family: 2,
		size: 10,
	},
	alignment: {
		vertical: "middle",
		horizontal: "center",
		// wrapText: true,
	},
};
export const borderStyle = {
	top_bottom: {
		border: { top: { style: "medium" }, bottom: { style: "medium" } },
	},
	left: { border: { left: { style: "medium" } } },
	right: { border: { right: { style: "medium" } } },
	bottom: { border: { bottom: { style: "medium" } } },
	side: { border: { right: { style: "medium" }, left: { style: "medium" } } },
	left_side: {
		border: {
			left: { style: "medium" },
			top: { style: "medium" },
			bottom: { style: "medium" },
		},
	},
	right_side: {
		border: {
			right: { style: "medium" },
			top: { style: "medium" },
			bottom: { style: "medium" },
		},
	},
};
