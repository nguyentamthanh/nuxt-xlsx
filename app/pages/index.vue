<template>
  <div class="min-h-screen flex items-center justify-center">
    <h1 class="text-white text-4xl font-bold">Hello World</h1>
  </div>
</template>

<script setup lang="ts">
import type { IConvert, IJira } from "@/types/common.type";

const filejira = ref<IJira[]>([]);
const fileconvert = ref<IConvert[]>([]);
const filerelease = ref<IConvert[]>([]);
onMounted(async () => {
  try {
    await Promise.all([readfileJira(), readfileconvert()]);
  } catch (error) {
    console.error("Error reading Excel file:", error);
  }
});
async function readfileJira(): Promise<void> {
  try {
    const file = await fetch("/files/jira.xlsx");
    const arrayBuffer = await file.arrayBuffer();
    const XLSX = await import("xlsx");
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName ?? ""];
    const json: IJira[] = XLSX.utils.sheet_to_json(sheet ?? {}) || [];
    filejira.value = json;
  } catch (error) {
    console.error("Error reading Excel file:", error);
  }
}
async function readfileconvert(): Promise<void> {
  try {
    const file = await fetch("/files/convert.xlsx");
    const arrayBuffer = await file.arrayBuffer();
    const XLSX = await import("xlsx");
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName ?? ""];
    const json: IConvert[] = XLSX.utils.sheet_to_json(sheet ?? {}) || [];
    fileconvert.value = json;
  } catch (error) {
    console.error("Error reading Excel file:", error);
  }
}
</script>
