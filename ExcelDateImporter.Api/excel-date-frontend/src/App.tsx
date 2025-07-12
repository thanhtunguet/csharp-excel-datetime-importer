import React, { useState } from "react";
import { Upload, Card, Table, message, Button, Divider } from "antd";
import { InboxOutlined } from "@ant-design/icons";
import Editor from "@monaco-editor/react";
import axios from "axios";
import type { UploadProps } from "antd";

const { Dragger } = Upload;

interface ExcelDateEntry {
  id: number;
  originalValue: string;
  parsedDate: string | null;
  dateFormat: string;
  isSuccessfullyParsed: boolean;
  errorMessage: string;
  createdAt: string;
}

function App() {
  const [parsedData, setParsedData] = useState<ExcelDateEntry[]>([]);
  const [loading, setLoading] = useState(false);
  const [csharpCode, setCsharpCode] = useState<string>("");

  const API_BASE_URL =
    import.meta.env.VITE_API_BASE_URL || window.location.origin;

  const fetchCodeExample = async () => {
    try {
      const response = await axios.get(
        `${API_BASE_URL}/api/excel/code-example`
      );
      setCsharpCode(response.data);
    } catch (error) {
      console.error("Error fetching code example:", error);
      message.error("Failed to fetch C# code example");
    }
  };

  React.useEffect(() => {
    fetchCodeExample();
  }, []);

  const uploadProps: UploadProps = {
    name: "file",
    multiple: false,
    accept: ".xlsx,.xls",
    customRequest: async ({ file, onSuccess, onError }) => {
      setLoading(true);
      const formData = new FormData();
      formData.append("file", file as File);

      try {
        const response = await axios.post(
          `${API_BASE_URL}/api/excel/upload`,
          formData,
          {
            headers: {
              "Content-Type": "multipart/form-data",
            },
          }
        );

        setParsedData(response.data);
        message.success("Excel file processed successfully!");
        onSuccess?.(response.data);
      } catch (error) {
        console.error("Upload error:", error);
        message.error("Failed to process Excel file");
        onError?.(error as Error);
      } finally {
        setLoading(false);
      }
    },
    onDrop(e) {
      console.log("Dropped files", e.dataTransfer.files);
    },
  };

  const columns = [
    {
      title: "Original Value",
      dataIndex: "originalValue",
      key: "originalValue",
      width: 150,
    },
    {
      title: "Parsed Date",
      dataIndex: "parsedDate",
      key: "parsedDate",
      width: 150,
      render: (date: string | null) =>
        date ? new Date(date).toLocaleDateString() : "N/A",
    },
    {
      title: "Date Format",
      dataIndex: "dateFormat",
      key: "dateFormat",
      width: 150,
    },
    {
      title: "Status",
      dataIndex: "isSuccessfullyParsed",
      key: "status",
      width: 100,
      render: (success: boolean) => (
        <span className={success ? "text-green-600" : "text-red-600"}>
          {success ? "Success" : "Failed"}
        </span>
      ),
    },
    {
      title: "Error Message",
      dataIndex: "errorMessage",
      key: "errorMessage",
      width: 200,
    },
  ];

  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-7xl mx-auto">
        <h1 className="text-3xl font-bold text-center mb-8 text-gray-800">
          Excel Date Format Importer Demo
        </h1>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Upload Section */}
          <Card title="Upload Excel File" className="shadow-lg">
            <Dragger {...uploadProps} className="mb-4">
              <p className="ant-upload-drag-icon">
                <InboxOutlined />
              </p>
              <p className="ant-upload-text">
                Click or drag Excel file to this area to upload
              </p>
              <p className="ant-upload-hint">
                Support for .xlsx and .xls files. The system will parse various
                date formats.
              </p>
            </Dragger>

            {parsedData.length > 0 && (
              <div className="mt-6">
                <h3 className="text-lg font-semibold mb-3">Parsed Results</h3>
                <Table
                  columns={columns}
                  dataSource={parsedData}
                  rowKey="id"
                  size="small"
                  scroll={{ x: 800, y: 400 }}
                  loading={loading}
                />
              </div>
            )}
          </Card>

          {/* Code Example Section */}
          <Card title="C# Code Example" className="shadow-lg">
            <p className="mb-4 text-gray-600">
              This C# code demonstrates how to handle various Excel date formats
              including serial dates, DateTime objects, and text dates in
              multiple formats.
            </p>
            <Button onClick={fetchCodeExample} className="mb-4" type="primary">
              Refresh Code Example
            </Button>
            <div className="border rounded-lg overflow-hidden">
              <Editor
                height="600px"
                defaultLanguage="csharp"
                value={csharpCode}
                theme="vs-dark"
                options={{
                  readOnly: true,
                  minimap: { enabled: false },
                  scrollBeyondLastLine: false,
                  fontSize: 12,
                }}
              />
            </div>
          </Card>
        </div>

        <Divider />

        <Card title="Supported Date Formats" className="shadow-lg">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div>
              <h4 className="font-semibold mb-2">Standard Formats</h4>
              <ul className="text-sm text-gray-600 space-y-1">
                <li>• dd/MM/yyyy (25/12/2023)</li>
                <li>• MM/dd/yyyy (12/25/2023)</li>
                <li>• yyyy-MM-dd (2023-12-25)</li>
                <li>• dd-MM-yyyy (25-12-2023)</li>
              </ul>
            </div>
            <div>
              <h4 className="font-semibold mb-2">Excel Specific</h4>
              <ul className="text-sm text-gray-600 space-y-1">
                <li>• Excel DateTime objects</li>
                <li>• Excel serial dates (44561)</li>
                <li>• dd.MM.yyyy (25.12.2023)</li>
                <li>• MM.dd.yyyy (12.25.2023)</li>
              </ul>
            </div>
            <div>
              <h4 className="font-semibold mb-2">Text Formats</h4>
              <ul className="text-sm text-gray-600 space-y-1">
                <li>• dd MMM yyyy (25 Dec 2023)</li>
                <li>• MMM dd, yyyy (Dec 25, 2023)</li>
                <li>• d/M/yyyy (5/3/2023)</li>
                <li>• M/d/yyyy (3/5/2023)</li>
              </ul>
            </div>
          </div>
        </Card>
      </div>
    </div>
  );
}

export default App;
