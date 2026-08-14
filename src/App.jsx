import { useRef, useState } from "react";

export default function App() {
  const inputRef = useRef(null);
  const pyodideRef = useRef(null);
  const [selectedFunction, setSelectedFunction] = useState("");
  const [sheetName, setSheetName] = useState("");

  const [file, setFile] = useState(null);
  const [status, setStatus] = useState("idle");
  // idle | processing | completed | error

  const [resultFile, setResultFile] = useState(null);

  const handleSelectFile = (selectedFile) => {
    if (!selectedFile) return;

    const fileName = selectedFile.name.toLowerCase();

    const isExcel = fileName.endsWith(".xlsx") || fileName.endsWith(".xlsm");

    if (!isExcel) {
      alert("Vui lòng chọn file Excel (.xlsx hoặc .xlsm).");
      return;
    }

    setFile(selectedFile);
    setResultFile(null);
    setStatus("idle");
  };

  const getPyodide = async () => {
    if (pyodideRef.current) {
      return pyodideRef.current;
    }

    const pyodide = await window.loadPyodide();

    // Cài micropip
    await pyodide.loadPackage("micropip");

    // Cài openpyxl vì script của bạn đang dùng openpyxl
    const micropip = pyodide.pyimport("micropip");

    await micropip.install("openpyxl");

    micropip.destroy();

    pyodideRef.current = pyodide;

    return pyodide;
  };

  const handleProcess = async () => {
    if (!selectedFunction) {
      alert("Vui lòng chọn chức năng xử lý.");
      return;
    }

    if (!file) {
      alert("Vui lòng chọn file Excel.");
      return;
    }
    if (!sheetName) {
      alert("Vui lòng nhập sheet.");
      return;
    }

    // Hiện tại mới tích hợp script xử lý chiết khấu
    if (selectedFunction !== "discount") {
      alert("Chức năng xử lý kho chưa được tích hợp.");
      return;
    }

    try {
      setStatus("processing");

      // ==========================
      // 1. Khởi tạo Pyodide
      // ==========================

      const pyodide = await getPyodide();

      // ==========================
      // 2. Xác định định dạng file
      // ==========================

      const isXlsm = file.name.toLowerCase().endsWith(".xlsm");

      const extension = isXlsm ? ".xlsm" : ".xlsx";

      const inputPath = `/input${extension}`;
      const outputPath = `/output${extension}`;

      const scriptPath = "/discount_processing.py";

      // ==========================
      // 3. Đưa file Excel vào
      //    filesystem của Pyodide
      // ==========================

      const buffer = await file.arrayBuffer();

      pyodide.FS.writeFile(inputPath, new Uint8Array(buffer));

      // Xóa output cũ nếu có
      try {
        pyodide.FS.unlink(outputPath);
      } catch {
        // File chưa tồn tại -> bỏ qua
      }

      // ==========================
      // 4. Load script Python
      // ==========================

      const scriptUrl = `${import.meta.env.BASE_URL}python/discount_processing.py`;

      const response = await fetch(scriptUrl);

      if (!response.ok) {
        throw new Error(`Không thể tải script Python: ${response.status}`);
      }

      const pythonCode = await response.text();

      pyodide.FS.writeFile(scriptPath, pythonCode, {
        encoding: "utf8",
      });

      // ==========================
      // 5. Tạo arguments
      //
      // python discount_processing.py
      // input.xlsx
      // output.xlsx
      // --sheet Sheet1
      // ==========================

      const args = [scriptPath, inputPath, outputPath];

      if (sheetName.trim()) {
        args.push("--sheet", sheetName.trim());
      }

      console.log("Python args:", args);

      // Chuyển JS array thành JSON
      const argsJson = JSON.stringify(args);

      // ==========================
      // 6. Chạy nguyên script Python
      // ==========================

      await pyodide.runPythonAsync(`
import sys
import runpy
import json

sys.argv = json.loads(${JSON.stringify(argsJson)})

runpy.run_path(
    ${JSON.stringify(scriptPath)},
    run_name="__main__"
)
    `);

      // ==========================
      // 7. Đọc file output
      // ==========================

      const outputData = pyodide.FS.readFile(outputPath);

      // ==========================
      // 8. Convert thành Blob
      // ==========================

      const mimeType = isXlsm
        ? "application/vnd.ms-excel.sheet.macroEnabled.12"
        : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

      const blob = new Blob([outputData], {
        type: mimeType,
      });

      const url = URL.createObjectURL(blob);

      // ==========================
      // 9. Hiển thị file download
      // ==========================

      setResultFile({
        name: `ket_qua_${file.name}`,
        url,
      });

      setStatus("completed");
    } catch (error) {
      console.error("Lỗi xử lý Python:", error);

      setStatus("error");
    }
  };

  const handleReset = () => {
    setFile(null);
    setResultFile(null);
    setSheetName("");
    setStatus("idle");

    if (inputRef.current) {
      inputRef.current.value = "";
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 p-8">
      <div className="mx-auto max-w-2xl">
        {/* Tiêu đề */}
        <div className="mb-6">
          <h1 className="text-2xl font-semibold text-gray-900">
            Công cụ xử lý file Excel
          </h1>

          <p className="mt-1 text-sm text-gray-500">
            Chọn chức năng, tải file Excel lên và thực hiện xử lý dữ liệu.
          </p>
        </div>

        <div className="rounded-xl border border-gray-200 bg-white p-6 shadow-sm">
          {/* Chọn chức năng */}
          <div className="mb-6">
            <p className="mb-2 text-sm font-medium text-gray-700">
              Chức năng xử lý
            </p>

            <div className="grid grid-cols-2 gap-3">
              {/* Xử lý chiết khấu */}
              <button
                type="button"
                disabled={status === "processing"}
                onClick={() => setSelectedFunction("discount")}
                className={`rounded-lg border p-4 text-left transition ${
                  selectedFunction === "discount"
                    ? "border-blue-500 bg-blue-50 ring-1 ring-blue-500"
                    : "border-gray-200 bg-white hover:border-blue-300 hover:bg-blue-50/30"
                }`}
              >
                <p className="text-sm font-medium text-gray-900">
                  Xử lý chiết khấu
                </p>

                <p className="mt-1 text-xs text-gray-500">
                  Xử lý và phân bổ dữ liệu chiết khấu hàng hóa
                </p>
              </button>

              {/* Xử lý kho */}
              <button
                type="button"
                disabled={status === "processing"}
                onClick={() => setSelectedFunction("inventory")}
                className={`rounded-lg border p-4 text-left transition ${
                  selectedFunction === "inventory"
                    ? "border-blue-500 bg-blue-50 ring-1 ring-blue-500"
                    : "border-gray-200 bg-white hover:border-blue-300 hover:bg-blue-50/30"
                }`}
              >
                <p className="text-sm font-medium text-gray-900">Xử lý kho</p>

                <p className="mt-1 text-xs text-gray-500">
                  Xử lý dữ liệu liên quan đến kho
                </p>
              </button>
            </div>
          </div>

          {/* Sheet */}
          <div className="mb-6">
            <label className="mb-2 block text-sm font-medium text-gray-700">
              Tên Sheet
            </label>

            <input
              type="text"
              value={sheetName}
              disabled={status === "processing"}
              onChange={(e) => setSheetName(e.target.value)}
              placeholder="Để trống nếu muốn xử lý tất cả các sheet"
              className="w-full rounded-lg border border-gray-300 px-3 py-2.5 text-sm text-gray-900 outline-none transition focus:border-blue-500 focus:ring-1 focus:ring-blue-500 disabled:bg-gray-100"
            />

            <p className="mt-1.5 text-xs text-gray-400">
              Không bắt buộc. Nhập tên sheet nếu chỉ muốn xử lý một sheet cụ
              thể.
            </p>
          </div>

          {/* Khu vực tải file */}
          {!file && (
            <div
              onClick={() => inputRef.current?.click()}
              className="flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed border-gray-300 px-6 py-14 transition hover:border-blue-400 hover:bg-blue-50/30"
            >
              <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-full bg-blue-50">
                <svg
                  className="h-6 w-6 text-blue-600"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                  strokeWidth="2"
                >
                  <path d="M12 16V4" />
                  <path d="m7 9 5-5 5 5" />
                  <path d="M5 20h14" />
                </svg>
              </div>

              <p className="text-sm font-medium text-gray-700">
                Nhấn để chọn file Excel
              </p>

              <p className="mt-1 text-xs text-gray-400">
                Định dạng hỗ trợ: .xlsx, .xlsm
              </p>

              <input
                ref={inputRef}
                type="file"
                accept=".xlsx,.xlsm"
                className="hidden"
                onChange={(e) => handleSelectFile(e.target.files?.[0])}
              />
            </div>
          )}

          {/* File đã chọn */}
          {file && (
            <div>
              <p className="mb-2 text-sm font-medium text-gray-700">
                File đầu vào
              </p>

              <div className="flex items-center gap-4 rounded-lg border border-gray-200 bg-gray-50 p-4">
                <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-lg bg-green-100">
                  <span className="text-sm font-bold text-green-700">XLS</span>
                </div>

                <div className="min-w-0 flex-1">
                  <p className="truncate text-sm font-medium text-gray-900">
                    {file.name}
                  </p>

                  <p className="mt-1 text-xs text-gray-500">
                    {(file.size / 1024).toFixed(1)} KB
                  </p>
                </div>

                {status === "idle" && (
                  <button
                    onClick={handleReset}
                    className="text-sm font-medium text-red-500 hover:text-red-600"
                  >
                    Xóa
                  </button>
                )}
              </div>
            </div>
          )}

          {/* Nút xử lý */}
          {file && status === "idle" && (
            <button
              onClick={handleProcess}
              disabled={!selectedFunction}
              className="mt-6 w-full rounded-lg bg-blue-600 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-blue-700 disabled:cursor-not-allowed disabled:bg-gray-300"
            >
              Xử lý file
            </button>
          )}

          {/* Đang xử lý */}
          {status === "processing" && (
            <div className="mt-6 rounded-lg border border-blue-200 bg-blue-50 p-4">
              <div className="flex items-center gap-3">
                <div className="h-5 w-5 animate-spin rounded-full border-2 border-blue-600 border-t-transparent" />

                <div>
                  <p className="text-sm font-medium text-blue-900">
                    Đang xử lý file...
                  </p>

                  <p className="mt-0.5 text-xs text-blue-600">
                    Vui lòng chờ trong khi hệ thống xử lý dữ liệu.
                  </p>
                </div>
              </div>
            </div>
          )}

          {/* Hoàn tất */}
          {status === "completed" && resultFile && (
            <div className="mt-6">
              <div className="mb-2 flex items-center gap-2">
                <div className="flex h-5 w-5 items-center justify-center rounded-full bg-green-100">
                  <span className="text-xs text-green-600">✓</span>
                </div>

                <p className="text-sm font-medium text-green-700">
                  Xử lý hoàn tất
                </p>
              </div>

              <div className="flex items-center gap-4 rounded-lg border border-green-200 bg-green-50 p-4">
                <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-lg bg-green-100">
                  <span className="text-sm font-bold text-green-700">XLS</span>
                </div>

                <div className="min-w-0 flex-1">
                  <p className="truncate text-sm font-medium text-gray-900">
                    {resultFile.name}
                  </p>

                  <p className="mt-1 text-xs text-gray-500">File đã xử lý</p>
                </div>

                <a
                  href={resultFile.url}
                  download={resultFile.name}
                  className="rounded-lg bg-green-600 px-4 py-2 text-sm font-medium text-white transition hover:bg-green-700"
                >
                  Tải xuống
                </a>
              </div>

              <button
                onClick={handleReset}
                className="mt-4 w-full rounded-lg border border-gray-300 px-4 py-2 text-sm font-medium text-gray-700 hover:bg-gray-50"
              >
                Xử lý file khác
              </button>
            </div>
          )}

          {/* Lỗi */}
          {status === "error" && (
            <div className="mt-6 rounded-lg border border-red-200 bg-red-50 p-4">
              <p className="text-sm font-medium text-red-700">
                Xử lý file thất bại.
              </p>

              <button
                onClick={() => setStatus("idle")}
                className="mt-2 text-sm font-medium text-red-600 underline"
              >
                Thử lại
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
