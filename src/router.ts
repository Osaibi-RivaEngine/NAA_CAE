import type { RouteRecordRaw } from "vue-router";
import HomeView from "./views/HomeView.vue";

export const routes: RouteRecordRaw[] = [
  {
    path: "/",
    name: "Home",
    component: HomeView,
  },
  {
    path: "/profile",
    name: "Profile",
    component: () => import("./views/ProfileView.vue"),
  },
];
