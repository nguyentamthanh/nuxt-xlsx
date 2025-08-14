export default defineNuxtConfig({
  compatibilityDate: '2025-07-15',
  devtools: { enabled: false },
  ssr: false,
  css: ["assets/css/main.css"],
  modules: ["@nuxt/ui", "dayjs-nuxt"],
});