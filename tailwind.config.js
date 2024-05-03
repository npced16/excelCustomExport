/** @type {import('tailwindcss').Config} */
module.exports = {
	// purge: [],
	purge: ['./index.html', './src/**/*.{vue,js,ts,jsx,tsx}'],
	darkMode: false, // or 'media' or 'class'
	mode: 'jit',
	theme: {
		extend: {
			flex: {
				1: '1',
				2: '2',
				2.5: '2.5',
				3: '3',
				4: '4',
				5: '5',
			},
			width: {
				'finedrop-card': '500px',
				'dashboard-card': '500px',
				inherit: 'inherit',
			},
			height: {
				inherit: 'inherit',
			},
			colors: {
				reservation: '#FFCB3B',
				processing: '#1C79D8',
				done: '#34BC9D',
				mci: '#BB4EFF',
				'mci-none': '#333333',
				'mci-prove': '#F29057',
				'mci-fail': '#EB5869',
				'mci-success': '#20B71A',
				'nurse-blue': '#6788FF',
				'button-blue': '#7499ff',
				'menu-blue': '#5f89ff',
				'gray-ab': '#ababab',
				'gray-6b': '#6b6b6b',
				'gray-f0': '#F0F0F0',
				'gray-9b': '#9b9b9b',
				'gray-150': '#b8c2cc',
				'gray-350': '#7c7c7c90',
				'black-31': '#313131',
				disabled: '#F5F5F5',
				danger: '#F25252',
				warning: '#FFCB3B',
				notice: '#4879FF',
			},
			boxShadow: {
				'card-shadow': ' 1px 1px 3px #242424a6',
			},
			fontFamily: {
				NotoSans: ['NotoSans'],
			},
			margin: {
				2.5: '10px',
			},
			gap: {
				2.5: '10px',
			},
		},
	},
	variants: {
		extend: {},
	},
	plugins: [],
};
