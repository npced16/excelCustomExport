import { createRouter, createWebHashHistory } from "vue-router";
import excelDownload from "@/components/examReservationExport.vue";
const routes = [
	{
		path: "/excelDownload",
		name: "Webpage",
		component: excelDownload,
	},
];

const router = createRouter({
	history: createWebHashHistory(process.env.BASE_URL),
	routes,
});

export default router;
