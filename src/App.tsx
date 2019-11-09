import React, { useEffect, useState, useCallback, useRef } from "react";
import XLSX, { WorkBook, WorkSheet } from "xlsx";
import "./App.css";
import { useDropzone } from "react-dropzone";
import Table from "react-bootstrap/Table";
import ListGroup from "react-bootstrap/ListGroup";
import Container from "react-bootstrap/Container";
import Col from "react-bootstrap/Col";
import Row from "react-bootstrap/Row";
import Button from "react-bootstrap/Button";

const readFile = (file: Blob): Promise<WorkBook> =>
  new Promise<WorkBook>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(f) {
      let data = new Uint8Array(f.target.result as ArrayBuffer);
      let workbook = XLSX.read(data, { type: "array" });
      resolve(workbook);
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });

const App: React.FC = () => {
  const [text, setText] = useState<string>("");
  const [id, setId] = useState<number>(0);
  const [file, setFile] = useState<File>();
  const [workbook, setWorkbook] = useState<WorkBook>(null);
  const [sheet, setSheet] = useState<WorkSheet>(null);
  const [chosenRows, setChosenRows] = useState<{
    [key: string]: boolean;
  }>(null);
  const [format, setFormat] = useState("");
  const outputRef = useRef(null);

  const onDrop = useCallback((acceptedFiles: File[]) => {
    setFile(acceptedFiles[0]);
  }, []);
  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    noClick: false,
    accept: [
      "application/vnd.ms-excel",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ]
  });

  const loadExcel = async (file: File) => {
    if (file) {
      const wb = await readFile(file);
      setWorkbook(wb);
      setFile(null);
    }
    if (workbook && Object.keys(workbook.Sheets).length === 1) {
      chooseWorksheet(Object.keys(workbook.Sheets)[0]);
    }
  };

  const chooseWorksheet = (workSheet: string) => {
    const chosenRows = {};
    const ws: any[] = XLSX.utils.sheet_to_json(workbook.Sheets[workSheet]);
    Object.keys(ws[0]).forEach(key => {
      chosenRows[key] = true;
    });
    setChosenRows(chosenRows);
    setSheet(ws.map(e => ({ ...e, selected: true })));
  };

  const addVariable = (variable: string) => {
    setFormat(format + `{${variable}}`);
  };

  const parseData = (): string[] => {
    const res = sheet
      .filter(e => e.selected)
      .map(row => {
        const regex = /\{(.*?)\}/g;
        return format.replace(regex, match => {
          const bracket = /\{(.*?)\}/g;
          const key = bracket.exec(match)[1];
          return row[key] || "-";
        });
      });
    return res;
  };

  const setAll = (selected: boolean) => {
    setSheet(sheet.map(e => ({ ...e, selected })));
  };

  const copyToClipboard = (e: React.MouseEvent) => {
    outputRef.current.select();
    document.execCommand("copy");
  };

  const selectRow = (row: number) => {
    const newSheet = sheet.map((e, idx) => {
      if (idx === row) {
        return { ...e, selected: !e.selected };
      } else return e;
    });
    setSheet(newSheet);
  };
  return (
    <div className="App">
      <Container>
        <header className="App-header">
          <h1>Easy Reg</h1>
          <h2>Easily Export your excel data in any format</h2>
          <div className="upload-zone" {...getRootProps()}>
            <input {...getInputProps()} />
            <p>Upload your file here</p>
          </div>
        </header>
        {file && (
          <>
            <p>{file.name}</p>
            <button onClick={() => loadExcel(file)}>Read Excel!</button>
          </>
        )}
      </Container>
      <Container>
        <Row>
          {workbook && (
            <Col md={sheet ? 4 : 12}>
              <h3>Choose worksheet</h3>
              <ListGroup>
                {Object.keys(workbook.Sheets).map((sheet, idx) => (
                  <ListGroup.Item
                    action
                    key={`workbook-${idx}`}
                    onClick={e => chooseWorksheet(sheet)}
                  >
                    {sheet}
                  </ListGroup.Item>
                ))}
              </ListGroup>
              {sheet && (
                <div>
                  <h3>Available Variables</h3>
                  <ListGroup>
                    {Object.keys(sheet[0]).map((column, idx) => (
                      <ListGroup.Item
                        action
                        onClick={e => addVariable(column)}
                        key={`sample-header-${idx}`}
                      >
                        {column}
                      </ListGroup.Item>
                    ))}
                  </ListGroup>
                  <h4>Type Format, replace variable with {"{<variable>}"}</h4>
                  <textarea
                    value={format}
                    onChange={e => setFormat(e.target.value)}
                  ></textarea>
                  <h3>Result</h3>
                  <small>Click to copy</small>
                  <textarea
                    ref={outputRef}
                    value={parseData().join("\n")}
                    onClick={copyToClipboard}
                    readOnly
                  ></textarea>
                </div>
              )}
            </Col>
          )}
          {sheet && (
            <Col className="scroll-wrapper">
              <Button variant="primary" onClick={() => setAll(true)}>
                Select All
              </Button>
              <Button variant="danger" onClick={() => setAll(false)}>
                Deselect All
              </Button>
              <Table>
                <thead>
                  <tr>
                    {Object.keys(sheet[0]).map((header, hidx) => (
                      <th key={`h-${hidx}`}>
                        <input
                          type="checkbox"
                          checked={chosenRows[header]}
                          onChange={e =>
                            setChosenRows({
                              ...chosenRows,
                              [header]: !chosenRows[header]
                            })
                          }
                        />{" "}
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {sheet.map((sheet, idx) => (
                    <tr
                      key={`r-${idx}`}
                      className={sheet.selected ? "selected" : ""}
                      onClick={e => selectRow(idx)}
                    >
                      {Object.values(sheet).map((value, cidx) => (
                        <td key={`r-${idx}-c-${cidx}`}>{value}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </Table>
            </Col>
          )}
        </Row>
      </Container>
    </div>
  );
};

export default App;
