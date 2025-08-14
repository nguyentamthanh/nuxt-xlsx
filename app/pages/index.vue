<template>
  <div class="min-h-screen p-5">
    <div class="max-w-7xl mx-auto">
      <!-- Header -->
      <div class="text-center mb-8">
        <h1 class="text-3xl font-bold text-gray-800 mb-2">
          Excel Data Processor
        </h1>
        <p class="text-gray-600">
          Cập nhật dữ liệu chấm công từ Jira vào file Excel
        </p>
      </div>

      <!-- Control Panel -->
      <div class="bg-white rounded-lg shadow-md p-6 mb-6">
        <div class="flex flex-wrap items-center gap-4 mb-4">
          <div class="flex-1 min-w-[200px]">
            <label class="block text-sm font-medium text-gray-700 mb-2"
              >Chọn Sheet</label
            >
            <USelect
              v-model="sheetName"
              :items="listsheet"
              placeholder="Chọn sheet để xem"
              class="w-full"
            />
          </div>
          <div class="flex gap-3">
            <UButton
              @click="convertToNewData"
              color="primary"
              :loading="isProcessing"
            >
              <UIcon name="i-heroicons-arrow-path" class="w-4 h-4 mr-2" />
              Cập nhật dữ liệu
            </UButton>
            <UButton
              @click="downloadNewFile"
              color="success"
              :disabled="!hasUpdatedData"
            >
              <UIcon name="i-heroicons-arrow-down-tray" class="w-4 h-4 mr-2" />
              Tải file mới
            </UButton>
          </div>
        </div>

        <!-- Status Info -->
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div class="bg-blue-50 p-4 rounded-lg">
            <div class="text-sm text-blue-600 font-medium">Tổng nhân viên</div>
            <div class="text-2xl font-bold text-blue-700">
              {{ fileconvert.length }}
            </div>
          </div>
          <div class="bg-green-50 p-4 rounded-lg">
            <div class="text-sm text-green-600 font-medium">Dữ liệu Jira</div>
            <div class="text-2xl font-bold text-green-700">
              {{ filejira.length }}
            </div>
          </div>
          <div class="bg-purple-50 p-4 rounded-lg">
            <div class="text-sm text-purple-600 font-medium">Trạng thái</div>
            <div class="text-2xl font-bold text-purple-700">
              {{ hasUpdatedData ? "Đã cập nhật" : "Chưa cập nhật" }}
            </div>
          </div>
        </div>
      </div>

      <!-- Data Display -->
      <div class="bg-white rounded-lg shadow-md overflow-hidden">
        <div class="px-6 py-4 border-b border-gray-200">
          <h2 class="text-lg font-semibold text-gray-800">Dữ liệu chấm công</h2>
          <p class="text-sm text-gray-600">
            Chỉ cập nhật phần ngày (cột 01-31) từ dữ liệu Jira
          </p>
        </div>

        <!-- Table Container with horizontal scroll -->
        <div class="overflow-x-auto">
          <table class="min-w-full divide-y divide-gray-200">
            <thead class="bg-gray-50">
              <tr>
                <th
                  class="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider sticky left-0 bg-gray-50 z-10"
                >
                  Thông tin
                </th>
                <th
                  v-for="day in 31"
                  :key="day"
                  class="px-2 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider min-w-[60px]"
                >
                  {{ day < 10 ? `0${day}` : day }}
                </th>
                <th
                  class="px-3 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider"
                >
                  Tổng ngày
                </th>
                <th
                  class="px-3 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider"
                >
                  Cuối tuần
                </th>
              </tr>
            </thead>
            <tbody class="bg-white divide-y divide-gray-200">
              <tr
                v-for="(item, index) in fileconvert"
                :key="index"
                class="hover:bg-gray-50"
              >
                <!-- Sticky left column with employee info -->
                <td
                  class="sticky left-0 bg-white px-3 py-4 whitespace-nowrap z-10 border-r border-gray-200"
                >
                  <div class="flex flex-col">
                    <div class="text-sm font-medium text-gray-900">
                      {{ item.Name }}
                    </div>
                    <div class="text-sm text-gray-500">{{ item.Role }}</div>
                    <div class="text-xs text-blue-600 font-medium">
                      {{ item.Author }}
                    </div>
                  </div>
                </td>

                <!-- Day columns -->
                <td
                  v-for="day in 31"
                  :key="day"
                  class="px-2 py-4 text-center text-sm border-l border-gray-100"
                  :class="getDayCellClass((item as any)[day < 10 ? `0${day}` : day.toString()])"
                >
                  <span
                    v-if="(item as any)[day < 10 ? `0${day}` : day.toString()] === '-'"
                    class="text-gray-400"
                    >-</span
                  >
                  <span
                    v-else-if="(item as any)[day < 10 ? `0${day}` : day.toString()] > 0"
                    class="font-medium text-green-700"
                  >
                    {{ (item as any)[day < 10 ? `0${day}` : day.toString()] }}
                  </span>
                  <span v-else class="text-gray-300">0</span>
                </td>

                <!-- Total days column -->
                <td
                  class="px-3 py-4 text-center text-sm font-medium bg-gray-50"
                >
                  <span class="text-blue-600">{{
                    (item as any)["Số ngày chấm công"] || 0
                  }}</span>
                </td>

                <!-- Weekend work column -->
                <td
                  class="px-3 py-4 text-center text-sm font-medium bg-orange-50"
                >
                  <span class="text-orange-600">{{
                    (item as any)["Chấm công cuối tuần"] || 0
                  }}</span>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- Legend -->
      <div class="mt-6 bg-white rounded-lg shadow-md p-4">
        <h3 class="text-sm font-medium text-gray-700 mb-3">Chú thích:</h3>
        <div class="flex flex-wrap gap-4 text-sm">
          <div class="flex items-center gap-2">
            <div
              class="w-4 h-4 bg-green-100 border border-green-300 rounded"
            ></div>
            <span class="text-gray-600">Có dữ liệu chấm công</span>
          </div>
          <div class="flex items-center gap-2">
            <div
              class="w-4 h-4 bg-gray-100 border border-gray-300 rounded"
            ></div>
            <span class="text-gray-600">Không có dữ liệu (-)</span>
          </div>
          <div class="flex items-center gap-2">
            <div
              class="w-4 h-4 bg-blue-100 border border-blue-300 rounded"
            ></div>
            <span class="text-gray-600">Tổng số ngày chấm công</span>
          </div>
          <div class="flex items-center gap-2">
            <div
              class="w-4 h-4 bg-orange-100 border border-orange-300 rounded"
            ></div>
            <span class="text-gray-600">Chấm công cuối tuần</span>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import type { IConvert, IJira } from "@/types/common.type";
import dayjs from "dayjs";

const filejira = ref<IJira[]>([]);
const fileconvert = ref<IConvert[]>([]);
const filerelease = ref<IConvert[]>([]);
const listsheet = ref<any[]>([]);
const sheetName = ref<string>("");
const isProcessing = ref(false);
const hasUpdatedData = ref(false);

onMounted(async () => {
  try {
    await Promise.all([readfileJira(), readfileconvert()]);
    convertToNewData();
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
    console.log("🚀 ~ readfileJira ~ json:", json);
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
    listsheet.value = workbook.SheetNames;
    const sheetName = workbook.SheetNames[1];
    const sheet = workbook.Sheets[sheetName ?? ""];
    const json: IConvert[] = XLSX.utils.sheet_to_json(sheet ?? {}) || [];
    const newArray = json.map((item: any, index: number) => {
      const converted: any = Object.fromEntries(
        Object.entries(item).map(([key, value]: [string, any]) => {
          const trimmedKey = key.trim();
          return /^\d+$/.test(trimmedKey)
            ? [Number(trimmedKey), value]
            : [trimmedKey, value];
        })
      );
      return converted;
    });
    fileconvert.value = newArray as unknown as IConvert[];
  } catch (error) {
    console.error("Error reading Excel file:", error);
  }
}

function convertToNewData(): void {
  isProcessing.value = true;

  try {
    // Lọc dữ liệu Jira theo tháng 7 năm 2025
    const dataMonth = filejira.value.filter((item: IJira) => {
      const excelDate = item["Started day"];
      if (typeof excelDate === "number") {
        const millisecondsPerDay = 24 * 60 * 60 * 1000;
        const excelEpoch = new Date(1900, 0, 1);
        const date = new Date(
          excelEpoch.getTime() + (excelDate - 1) * millisecondsPerDay
        );
        const dayjsDate = dayjs(date);
        const month = dayjsDate.month() + 1;
        const year = dayjsDate.year();

        return month === 7 && year === 2025;
      }
      return false;
    });

    console.log("🚀 ~ dataMonth (tháng 7):", dataMonth);

    // Cập nhật dữ liệu convert với time spent từ Jira
    const updatedConvert = fileconvert.value.map((convertItem: IConvert) => {
      // Tìm tất cả dữ liệu Jira của author này
      const jiraItems = dataMonth.filter(
        (jiraItem: IJira) =>
          convertItem.Author?.trim() === jiraItem.Author?.trim()
      );

      if (jiraItems.length === 0) {
        // Nếu không có dữ liệu Jira, để nguyên dữ liệu cũ
        return convertItem;
      }

      // Tạo object mới với dữ liệu cũ
      const updatedItem: any = { ...convertItem };
      let soNgayChamCong = 0; // Đếm số ngày có chấm công
      let chamCongCuoiTuan = 0; // Tổng time spent cuối tuần

      // Cập nhật từng ngày trong tháng
      for (let day = 1; day <= 31; day++) {
        const dayKey = day < 10 ? `0${day}` : day.toString(); // Format: 01, 02, 03...

        // Tìm tất cả time spent của ngày này
        const dayItems = jiraItems.filter((jiraItem: IJira) => {
          const excelDate = jiraItem["Started day"];
          if (typeof excelDate === "number") {
            const millisecondsPerDay = 24 * 60 * 60 * 1000;
            const excelEpoch = new Date(1900, 0, 1);
            const date = new Date(
              excelEpoch.getTime() + (excelDate - 1) * millisecondsPerDay
            );
            const dayjsDate = dayjs(date);
            return dayjsDate.date() === day;
          }
          return false;
        });

        if (dayItems.length > 0) {
          // Cộng dồn time spent của ngày này
          let totalTimeSpent = 0;
          dayItems.forEach((jiraItem: IJira) => {
            const timeSpent = parseFloat(jiraItem["Time spent"] || "0");
            if (!isNaN(timeSpent)) {
              totalTimeSpent += timeSpent;
            }
          });

          // Cập nhật vào ngày tương ứng
          if (dayKey in updatedItem) {
            updatedItem[dayKey] = totalTimeSpent;
            // Tăng số ngày chấm công nếu có time spent > 0
            if (totalTimeSpent > 0) {
              soNgayChamCong++;

              // Kiểm tra nếu là cuối tuần (thứ 7 hoặc chủ nhật)
              const excelDate = dayItems[0]?.["Started day"];
              if (typeof excelDate === "number") {
                const millisecondsPerDay = 24 * 60 * 60 * 1000;
                const excelEpoch = new Date(1900, 0, 1);
                const date = new Date(
                  excelEpoch.getTime() + (excelDate - 1) * millisecondsPerDay
                );
                const dayOfWeek = date.getDay(); // 0 = Chủ nhật, 6 = Thứ 7
                if (dayOfWeek === 0 || dayOfWeek === 6) {
                  chamCongCuoiTuan += totalTimeSpent;
                }
              }
            }
          }
        } else {
          // Nếu không có dữ liệu, để dấu "-"
          if (dayKey in updatedItem) {
            updatedItem[dayKey] = "-";
          }
        }
      }

      // Cập nhật trường "Số ngày chấm công"
      if ("Số ngày chấm công" in updatedItem) {
        updatedItem["Số ngày chấm công"] = soNgayChamCong;
      }

      // Cập nhật trường "Chấm công cuối tuần"
      if ("Chấm công cuối tuần" in updatedItem) {
        updatedItem["Chấm công cuối tuần"] = chamCongCuoiTuan;
      }

      console.log(`🚀 ~ Updated ${convertItem.Author}:`, updatedItem);
      console.log(`🚀 ~ Số ngày chấm công: ${soNgayChamCong}`);
      console.log(`🚀 ~ Chấm công cuối tuần: ${chamCongCuoiTuan}`);
      return updatedItem as IConvert;
    });

    // Cập nhật state
    fileconvert.value = updatedConvert;
    hasUpdatedData.value = true;

    console.log("🚀 ~ Final updated convert data:", fileconvert.value);
  } catch (error) {
    console.error("Error in convertToNewData:", error);
  } finally {
    isProcessing.value = false;
  }
}

// Hàm helper để xác định class cho ô ngày
function getDayCellClass(value: any): string {
  if (value === "-") return "bg-gray-50";
  if (value > 0) return "bg-green-50";
  return "bg-white";
}

// Hàm tải file Excel mới xuống
async function downloadNewFile(): Promise<void> {
  try {
    const XLSX = await import("xlsx");

    // Đọc lại file convert gốc để lấy cấu trúc sheet
    const file = await fetch("/files/convert.xlsx");
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    // Lấy sheet gốc (giữ nguyên cấu trúc)
    const sheetName = workbook.SheetNames[1]; // Sheet thứ 2 như trong readfileconvert
    const originalSheet = workbook.Sheets[sheetName ?? ""];

    if (!originalSheet) {
      throw new Error("Không thể tìm thấy sheet gốc");
    }

    // Cập nhật từng ô một cách chính xác để giữ nguyên cấu trúc cột
    fileconvert.value.forEach((rowData: IConvert, rowIndex: number) => {
      const excelRow = rowIndex + 2; // Bắt đầu từ dòng 2 (sau header)

      // Cập nhật từng cột theo thứ tự gốc
      Object.entries(rowData).forEach(([key, value]) => {
        if (
          key === "No." ||
          key === "Name" ||
          key === "Role" ||
          key === "Author" ||
          key === "Số ngày chấm công" ||
          key === "Chấm công cuối tuần"
        ) {
          // Các cột text
          const colIndex = getColumnIndex(key);
          if (colIndex !== -1) {
            const cellAddress = XLSX.utils.encode_cell({
              r: excelRow - 1,
              c: colIndex,
            });
            originalSheet[cellAddress] = {
              v: value,
              t: typeof value === "number" ? "n" : "s",
            };
          }
        } else if (/^\d{2}$/.test(key)) {
          // Các cột ngày (01, 02, 03...)
          const colIndex = getColumnIndex(key);
          if (colIndex !== -1) {
            const cellAddress = XLSX.utils.encode_cell({
              r: excelRow - 1,
              c: colIndex,
            });
            originalSheet[cellAddress] = {
              v: value,
              t: typeof value === "number" ? "n" : "s",
            };
          }
        }
      });
    });

    // Cập nhật sheet trong workbook
    workbook.Sheets[sheetName ?? ""] = originalSheet;

    // Tạo file buffer với cấu trúc gốc
    const excelBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
    });

    // Tạo blob và download
    const blob = new Blob([excelBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);

    // Tạo link download
    const link = document.createElement("a");
    link.href = url;
    link.download = "new_update.xlsx";
    document.body.appendChild(link);
    link.click();

    // Dọn dẹp
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);

    console.log(
      "🚀 ~ File đã được tải xuống: new_update.xlsx (giữ nguyên cấu trúc cột gốc)"
    );
  } catch (error) {
    console.error("🚀 ~ Lỗi khi tải file:", error);
  }
}

// Hàm helper để lấy index cột từ tên cột
function getColumnIndex(columnName: string): number {
  const columnMap: { [key: string]: number } = {
    "No.": 0, // A - Số thứ tự
    Author: 1, // B - Tác giả
    Name: 2, // C - Tên
    Role: 3, // D - Vai trò
    "01": 4, // E - Ngày 01
    "02": 5, // F - Ngày 02
    "03": 6, // G - Ngày 03
    "04": 7, // H - Ngày 04
    "05": 8, // I - Ngày 05
    "06": 9, // J - Ngày 06
    "07": 10, // K - Ngày 07
    "08": 11, // L - Ngày 08
    "09": 12, // M - Ngày 09
    "10": 13, // N - Ngày 10
    "11": 14, // O - Ngày 11
    "12": 15, // P - Ngày 12
    "13": 16, // Q - Ngày 13
    "14": 17, // R - Ngày 14
    "15": 18, // S - Ngày 15
    "16": 19, // T - Ngày 16
    "17": 20, // U - Ngày 17
    "18": 21, // V - Ngày 18
    "19": 22, // W - Ngày 19
    "20": 23, // X - Ngày 20
    "21": 24, // Y - Ngày 21
    "22": 25, // Z - Ngày 22
    "23": 26, // AA - Ngày 23
    "24": 27, // AB - Ngày 24
    "25": 28, // AC - Ngày 25
    "26": 29, // AD - Ngày 26
    "27": 30, // AE - Ngày 27
    "28": 31, // AF - Ngày 28
    "29": 32, // AG - Ngày 29
    "30": 33, // AH - Ngày 30
    "31": 34, // AI - Ngày 31
    "Số ngày chấm công": 35, // AJ - Số ngày chấm công
    "Chấm công cuối tuần": 36, // AK - Chấm công cuối tuần
  };

  return columnMap[columnName] || -1;
}
</script>
