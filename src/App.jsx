import { useRef, useState } from "react";

export default function App() {
  const inputRef = useRef(null);

  const [file, setFile] = useState(null);
  const [status, setStatus] = useState("idle");
  // idle | processing | completed | error

  const [resultFile, setResultFile] = useState(null);

  const handleSelectFile = (selectedFile) => {
    if (!selectedFile) return;

    const isExcel =
      selectedFile.name.endsWith(".xlsx") || selectedFile.name.endsWith(".xls");

    if (!isExcel) {
      alert("Please select an Excel file.");
      return;
    }

    setFile(selectedFile);
    setResultFile(null);
    setStatus("idle");
  };

  const handleProcess = async () => {
    if (!file) return;

    try {
      setStatus("processing");

      // ==========================
      // MOCK API PROCESSING
      // Replace bằng API thật sau
      // ==========================
      await new Promise((resolve) => setTimeout(resolve, 2000));

      const buffer = await file.arrayBuffer();

      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      const url = URL.createObjectURL(blob);

      setResultFile({
        name: "processed.xlsx",
        url,
      });

      setStatus("completed");
    } catch (error) {
      console.error(error);
      setStatus("error");
    }
  };

  const handleReset = () => {
    setFile(null);
    setResultFile(null);
    setStatus("idle");

    if (inputRef.current) {
      inputRef.current.value = "";
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 p-8">
      <div className="mx-auto max-w-2xl">
        {/* Header */}
        <div className="mb-6">
          <h1 className="text-2xl font-semibold text-gray-900">
            Excel File Processing
          </h1>

          <p className="mt-1 text-sm text-gray-500">
            Upload an Excel file, process it, and download the result.
          </p>
        </div>

        <div className="rounded-xl border border-gray-200 bg-white p-6 shadow-sm">
          {/* Upload area */}
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
                Click to upload Excel file
              </p>

              <p className="mt-1 text-xs text-gray-400">
                Supported formats: .xlsx, .xls
              </p>

              <input
                ref={inputRef}
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={(e) => handleSelectFile(e.target.files?.[0])}
              />
            </div>
          )}

          {/* Selected file */}
          {file && (
            <div>
              <p className="mb-2 text-sm font-medium text-gray-700">
                Input File
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
                    Remove
                  </button>
                )}
              </div>
            </div>
          )}

          {/* Process button */}
          {file && status === "idle" && (
            <button
              onClick={handleProcess}
              className="mt-6 w-full rounded-lg bg-blue-600 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-blue-700"
            >
              Process File
            </button>
          )}

          {/* Processing */}
          {status === "processing" && (
            <div className="mt-6 rounded-lg border border-blue-200 bg-blue-50 p-4">
              <div className="flex items-center gap-3">
                <div className="h-5 w-5 animate-spin rounded-full border-2 border-blue-600 border-t-transparent" />

                <div>
                  <p className="text-sm font-medium text-blue-900">
                    Processing file...
                  </p>

                  <p className="mt-0.5 text-xs text-blue-600">
                    Please wait while the file is being processed.
                  </p>
                </div>
              </div>
            </div>
          )}

          {/* Completed */}
          {status === "completed" && resultFile && (
            <div className="mt-6">
              <div className="mb-2 flex items-center gap-2">
                <div className="flex h-5 w-5 items-center justify-center rounded-full bg-green-100">
                  <span className="text-xs text-green-600">✓</span>
                </div>

                <p className="text-sm font-medium text-green-700">
                  Processing completed
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

                  <p className="mt-1 text-xs text-gray-500">Processed file</p>
                </div>

                <a
                  href={resultFile.url}
                  download={resultFile.name}
                  className="rounded-lg bg-green-600 px-4 py-2 text-sm font-medium text-white transition hover:bg-green-700"
                >
                  Download
                </a>
              </div>

              <button
                onClick={handleReset}
                className="mt-4 w-full rounded-lg border border-gray-300 px-4 py-2 text-sm font-medium text-gray-700 hover:bg-gray-50"
              >
                Process Another File
              </button>
            </div>
          )}

          {/* Error */}
          {status === "error" && (
            <div className="mt-6 rounded-lg border border-red-200 bg-red-50 p-4">
              <p className="text-sm font-medium text-red-700">
                Failed to process file.
              </p>

              <button
                onClick={() => setStatus("idle")}
                className="mt-2 text-sm font-medium text-red-600 underline"
              >
                Try again
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
