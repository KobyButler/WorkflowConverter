document.addEventListener('DOMContentLoaded', function () {
    const excelFileInput = document.getElementById('excelFileInput');
    const processButton = document.getElementById('processButton');
    const statusMessage = document.getElementById('statusMessage');
    const fullJSON = document.getElementById('fullJSON');
  
    processButton.addEventListener('click', async () => {
      if (!excelFileInput.files || excelFileInput.files.length === 0) {
        statusMessage.textContent = 'Please select an Excel file.';
        return;
      }
  
      const selectedFile = excelFileInput.files[0];
      const reader = new FileReader();
  
      reader.onload = async (event) => {
        const arrayBuffer = event.target.result;
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
  
        const allProcessedData = [];
  
        let fid = 100;
  
        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
  
          const processedData = [];
  
          const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  
          let id = 1;
          let labelAfter = 1;
  
          let outputContent = `{\n`;
          outputContent += `\t"id": ${id},\n`;
          outputContent += `\t"key": "F${fid}",\n`;
          outputContent += `\t"name": "${sheetName}",\n`;
          outputContent += `\t"groups": [\n`;
          fid++;
  
          allRows.forEach((row, idx) => {
            //inputColumns[0] == Label
            //inputColumns[1] == Items within Label
            //inputColumns[2] == Type
            //inputColumns[3] == Edit Notes
            //inputColumns[4] == In/Out Controls
            //inputColumns[5] == Required?
            //inputColumns[6] == Conditional
            //inputColumns[7] == Triggering Field
            //inputColumns[8] == Triggering Item/Value
            //inputColumns[9] == id

            let inputColumns = row.slice(0, 9);
            let firstItem = 1
  
            if (idx === 0) {
              return; // Skip header row
            }
  
            if (inputColumns.every(cell => cell === undefined)) {
              return; // Skip empty row
            }
  
            if (inputColumns[0] === undefined) {
              return;
            }
  
            let previousRow = processedData.length > 0 ? processedData[processedData.length - 1]['current'] : null;
            let nextRow = idx < allRows.length - 1 ? allRows[idx + 1].slice(0, 3) : null;
  
            let comparison = {
              'current': inputColumns,
              'previous': [],
              'next': []
            };
  
            if (previousRow) {
              comparison['previous'] = previousRow.map((previousValue, i) =>
                `Prev Column ${i + 1}: ${previousValue} (Current: ${inputColumns[i]})`
              );
            }
  
            if (nextRow) {
              comparison['next'] = nextRow.map((nextValue, i) =>
                `Next Column ${i + 1}: ${nextValue} (Current: ${inputColumns[i]})`
              );
            }
  
            processedData.push(comparison);
            labelAfter = 1;
            let needNewLine = 1;
  
            outputContent += `\t\t{\n`;
          //  outputContent += `\t\t\t"id": "${inputColumns[3]}",\n`;
            outputContent += `\t\t\t"id": ${id},\n`;
  
            if (inputColumns[1] === undefined) {
              outputContent += `\t\t\t"label": "${inputColumns[0]}",\n`;
              outputContent += `\t\t\t"type": "${inputColumns[2]}"`;
              
          /*    
              if (inputColumns[5] === true) {
                outputContent += `,\n`;
                let strid = inputColumns[6];
                let newStrid = strid.replace(/\s/g, ' ');
                outputContent += `\t\t\t"conditional" "${newStrid}"`;
              }
          */
              if (inputColumns[5] === true) {
                outputContent += `,\n`;
                outputContent += `\t\t\t"required": true`;
              }

              if (inputColumns[4] === true) {
                outputContent += `,\n`;
                outputContent += `\t\t\t"pre_post": true,\n`;
                outputContent += `\t\t\t"pre_label": "Out",\n`;
                outputContent += `\t\t\t"post_label": "In"\n`;
              }
              else {
                outputContent += `\n`;
              }

              labelAfter = 0;
            } else {
              outputContent += `\t\t\t"type": "${inputColumns[2]}",\n`;
              let multiLabel = inputColumns[0];
              if (inputColumns[1] !== undefined) {
                id++;
                outputContent += `\t\t\t"items": [\n`;
                outputContent += `\t\t\t\t{\n`;
                outputContent += `\t\t\t\t\t"id": ${id},\n`;
                outputContent += `\t\t\t\t\t"title": "${inputColumns[1]}"\n`;
                outputContent += `\t\t\t\t}`;
                id++;
                needNewLine = 1
  
                while ((nextRow[0] === undefined || nextRow[0] === inputColumns[0]) && nextRow[1] !== undefined) {
                    needNewLine = 0;

                  if (firstItem === 1) {
                    outputContent += `,\n`;
                    firstItem = 0;
                  }
                  outputContent += `\t\t\t\t{\n`;
                  outputContent += `\t\t\t\t\t"id": ${id},\n`;
                  outputContent += `\t\t\t\t\t"title": "${nextRow[1]}"\n`;
                  current = nextRow;
  
                  id++;
                  idx++;
                  nextRow = idx < allRows.length - 1 ? allRows[idx + 1].slice(0, 3) : null;
  
                  if ((nextRow[2] === current[2]) && ((nextRow[0] === multiLabel) || (nextRow[0] === undefined))) {
                    outputContent += `\t\t\t\t},\n`;
                  } else if (nextRow[2] !== undefined) {
                    outputContent += `\t\t\t\t}\n`;
                  } else if (current === undefined) {
                    outputContent += `\t\t\t\t}\n`;
                  } else {
                    outputContent += `\t\t\t\t,\n`;
                  }
                }
                if (needNewLine === 1)
                    outputContent += `\n`;

                outputContent += `\t\t\t],\n`;
                if (labelAfter === 1) {
                  outputContent += `\t\t\t"label": "${multiLabel}"`;
                  
              /*    if (inputColumns[5] === true) {
                    outputContent += `,\n`;
                    let strid = inputColumns[6];
                    let newStrid = strid.replace(/\s/g, ' ');
                    outputContent += `\t\t\t"conditional" "${newStrid}"`;
                  }
              */
                  if (inputColumns[5] === true) {
                    outputContent += `,\n`;
                    outputContent += `\t\t\t"required": true`;
                  }

                  if (inputColumns[4] === true) {
                    outputContent += `,\n`;
                    outputContent += `\t\t\t"pre_post": true,\n`;
                    outputContent += `\t\t\t"pre_label": "Out",\n`;
                    outputContent += `\t\t\t"post_label": "In"\n`;
                  }
                  else {
                    outputContent += `\n`;
                  }

                  labelAfter = 0;
                }
              }
            }
            outputContent += `\t\t},\n`;
            id++;
          });
  
          outputContent += `\t],\n`;
          outputContent += `\t"tab_bar_item": {\n`;
          outputContent += `\t\t"id": "",\n`;
          outputContent += `\t\t"url": "",\n`;
          outputContent += `\t\t"title": "${sheetName}",\n`;
          outputContent += `\t\t"image_name": "formDisabled.png"\n`;
          outputContent += `\t}\n`;
          outputContent += `},\n`;
  
          outputContent = outputContent.replace(/,\n\t],/g, '\n\t],');
          outputContent = outputContent.replace(/,\n\t\t\t],/g, '\n\t\t\t],');
  
          const outputFilePath = `${sheetName}_${'output'}.txt`;
          
          const blob = new Blob([outputContent], { type: 'text/plain' });
  
          const a = document.createElement('a');
          a.href = URL.createObjectURL(blob);
          a.download = outputFilePath;
          a.click();
  
          URL.revokeObjectURL(a.href);
  
          statusMessage.textContent = 'File processed and downloaded successfully.';
  
          allProcessedData.push([sheetName, outputContent]);
        });
  
        const combinedOutputFilePath = `${'combined'}.txt`;
  
        const combinedOutputContent = allProcessedData.map(([_, sheetContent]) => sheetContent).join('');
  
        const combinedBlob = new Blob([combinedOutputContent], { type: 'text/plain' });
  
        const aCombined = document.createElement('a');
        aCombined.href = URL.createObjectURL(combinedBlob);
        aCombined.download = combinedOutputFilePath;
        aCombined.click();
  
        URL.revokeObjectURL(aCombined.href);
      };
  
      reader.onerror = (event) => {
        statusMessage.textContent = 'An error occurred while reading the file.';
      };
  
      reader.readAsArrayBuffer(selectedFile);
    });

    fullJSON.addEventListener('click', async () => {
      if (!excelFileInput.files || excelFileInput.files.length === 0) {
        statusMessage.textContent = 'Please select an Excel file.';
        return;
      }

      let outputContent = [
        `{
          "template": {
            "id": 1,
            "modules": [
              {
                "id": 2,
                "key": "RC2",
                "home": true,
                "name": "Reference Number Capture 2",
                "reference_entry": {
                  "id": 3,
                  "prompt": "Enter Unit or Reference Number"
                },
                "keyboard_autocapitalization_type": "all_letters"
              },
              {
                "id": 4,
                "key": "M1",
                "name": "Media Capture Module 1",
                "instructions": [
                  {
                    "id": 5,
                    "label": "Take photo of fuel level and mileage on dash",
                    "autonav": true,
                    "auto_advance": true
                  }
                ],
                "tab_bar_item": {
                  "id": 6,
                  "url": "",
                  "title": "Record360",
                  "image_name": "R360icon.png"
                }
              },
              {
                "id": 7,
                "key": "N2",
                "name": "Notations Module 2",
                "notations": [
                  {
                    "id": 8,
                    "rank": 1,
                    "title": "Broken",
                    "autonav": true
                  },
                  {
                    "id": 9,
                    "rank": 2,
                    "title": "Damage",
                    "autonav": true
                  },
                  {
                    "id": 10,
                    "rank": 3,
                    "title": "Dent",
                    "autonav": true
                  },
                  {
                    "id": 11,
                    "rank": 4,
                    "title": "Hole",
                    "autonav": true
                  },
                  {
                    "id": 12,
                    "rank": 5,
                    "title": "Missing",
                    "autonav": true
                  },
                  {
                    "id": 13,
                    "rank": 6,
                    "title": "Paint Chip",
                    "autonav": true
                  },
                  {
                    "id": 14,
                    "rank": 7,
                    "title": "Prior",
                    "autonav": true
                  },
                  {
                    "id": 15,
                    "rank": 8,
                    "title": "Torn",
                    "autonav": true
                  },
                  {
                    "id": 16,
                    "rank": 9,
                    "title": "Other",
                    "autonav": false
                  }
                ],
                "tab_bar_item": {
                  "id": 17,
                  "url": "",
                  "title": "Notations",
                  "image_name": "notation.png"
                },
                "allowmultiselect": true
              },
              {
                "id": 18,
                "key": "LB1",
                "name": "Logic Branch 1",
                "logic_branch": {
                  "id": 19,
                  "prompt": "You have entered an existing Unit or Reference Number. Do you want to Update this record, apply a Return to this record or start a new Checkout record?",
                  "values": [
                    {
                      "id": 20,
                      "value": "true",
                      "action": "DD1"
                    },
                    {
                      "id": 21,
                      "value": "false",
                      "action": "M1"
                    }
                  ],
                  "options": [
                    {
                      "id": 22,
                      "label": "Update",
                      "value": null
                    },
                    {
                      "id": 23,
                      "label": "Return",
                      "value": true
                    },
                    {
                      "id": 24,
                      "label": "Checkout",
                      "value": false
                    }
                  ],
                  "variable": "{RETURN}"
                }
              },
              {
                "id": 25,
                "key": "LB2",
                "name": "Logic Branch 2",
                "logic_branch": {
                  "id": 26,
                  "values": [
                    {
                      "id": 27,
                      "value": "false",
                      "action": "DD2"
                    }
                  ],
                  "variable": "{DAMAGE}"
                }
              },
        `
      ];

      let allProcessedData = [];
  
      allProcessedData.push([outputContent]);

      const selectedFile = excelFileInput.files[0];
      const reader = new FileReader();
  
      reader.onload = async (event) => {
        const arrayBuffer = event.target.result;
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
  
        let lb6 = [];
  
        let fid = 100;
  
        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
  
          const processedData = [];
  
          const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  
          let id = 28;
          let labelAfter = 1;
  
          outputContent = `{\n`;
          outputContent += `\t"id": ${id},\n`;
          outputContent += `\t"key": "F${fid}",\n`;
          outputContent += `\t"name": "${sheetName}",\n`;
          outputContent += `\t"groups": [\n`;
          fid++;
  
          allRows.forEach((row, idx) => {
            //inputColumns[0] == Label
            //inputColumns[1] == Items within Label
            //inputColumns[2] == Type
            //inputColumns[3] == Edit Notes
            //inputColumns[4] == In/Out Controls
            //inputColumns[5] == Required?
            //inputColumns[6] == Conditional
            //inputColumns[7] == Triggering Field
            //inputColumns[8] == Triggering Item/Value
            //inputColumns[9] == id

            let inputColumns = row.slice(0, 9);
            let firstItem = 1
  
            if (idx === 0) {
              return; // Skip header row
            }
  
            if (inputColumns.every(cell => cell === undefined)) {
              return; // Skip empty row
            }
  
            if (inputColumns[0] === undefined) {
              return;
            }
  
            let previousRow = processedData.length > 0 ? processedData[processedData.length - 1]['current'] : null;
            let nextRow = idx < allRows.length - 1 ? allRows[idx + 1].slice(0, 3) : null;
  
            let comparison = {
              'current': inputColumns,
              'previous': [],
              'next': []
            };
  
            if (previousRow) {
              comparison['previous'] = previousRow.map((previousValue, i) =>
                `Prev Column ${i + 1}: ${previousValue} (Current: ${inputColumns[i]})`
              );
            }
  
            if (nextRow) {
              comparison['next'] = nextRow.map((nextValue, i) =>
                `Next Column ${i + 1}: ${nextValue} (Current: ${inputColumns[i]})`
              );
            }
  
            processedData.push(comparison);
            labelAfter = 1;
            let needNewLine = 1;
  
            outputContent += `\t\t{\n`;
          //  outputContent += `\t\t\t"id": "${inputColumns[3]}",\n`;
            outputContent += `\t\t\t"id": ${id},\n`;
  
            if (inputColumns[1] === undefined) {
              outputContent += `\t\t\t"label": "${inputColumns[0]}",\n`;
              outputContent += `\t\t\t"type": "${inputColumns[2]}"`;
              
          /*    
              if (inputColumns[5] === true) {
                outputContent += `,\n`;
                let strid = inputColumns[6];
                let newStrid = strid.replace(/\s/g, ' ');
                outputContent += `\t\t\t"conditional" "${newStrid}"`;
              }
          */
              if (inputColumns[5] === true) {
                outputContent += `,\n`;
                outputContent += `\t\t\t"required": true`;
              }

              if (inputColumns[4] === true) {
                outputContent += `,\n`;
                outputContent += `\t\t\t"pre_post": true,\n`;
                outputContent += `\t\t\t"pre_label": "Out",\n`;
                outputContent += `\t\t\t"post_label": "In"\n`;
              }
              else {
                outputContent += `\n`;
              }

              labelAfter = 0;
            } else {
              outputContent += `\t\t\t"type": "${inputColumns[2]}",\n`;
              let multiLabel = inputColumns[0];
              if (inputColumns[1] !== undefined) {
                id++;
                outputContent += `\t\t\t"items": [\n`;
                outputContent += `\t\t\t\t{\n`;
                outputContent += `\t\t\t\t\t"id": ${id},\n`;
                outputContent += `\t\t\t\t\t"title": "${inputColumns[1]}"\n`;
                outputContent += `\t\t\t\t}`;
                id++;
                needNewLine = 1
  
                while ((nextRow[0] === undefined || nextRow[0] === inputColumns[0]) && nextRow[1] !== undefined) {
                    needNewLine = 0;

                  if (firstItem === 1) {
                    outputContent += `,\n`;
                    firstItem = 0;
                  }
                  outputContent += `\t\t\t\t{\n`;
                  outputContent += `\t\t\t\t\t"id": ${id},\n`;
                  outputContent += `\t\t\t\t\t"title": "${nextRow[1]}"\n`;
                  current = nextRow;
  
                  id++;
                  idx++;
                  nextRow = idx < allRows.length - 1 ? allRows[idx + 1].slice(0, 3) : null;
  
                  if ((nextRow[2] === current[2]) && ((nextRow[0] === multiLabel) || (nextRow[0] === undefined))) {
                    outputContent += `\t\t\t\t},\n`;
                  } else if (nextRow[2] !== undefined) {
                    outputContent += `\t\t\t\t}\n`;
                  } else if (current === undefined) {
                    outputContent += `\t\t\t\t}\n`;
                  } else {
                    outputContent += `\t\t\t\t,\n`;
                  }
                }
                if (needNewLine === 1)
                    outputContent += `\n`;

                outputContent += `\t\t\t],\n`;
                if (labelAfter === 1) {
                  outputContent += `\t\t\t"label": "${multiLabel}"`;
                  
              /*    if (inputColumns[5] === true) {
                    outputContent += `,\n`;
                    let strid = inputColumns[6];
                    let newStrid = strid.replace(/\s/g, ' ');
                    outputContent += `\t\t\t"conditional" "${newStrid}"`;
                  }
              */
                  if (inputColumns[5] === true) {
                    outputContent += `,\n`;
                    outputContent += `\t\t\t"required": true`;
                  }

                  if (inputColumns[4] === true) {
                    outputContent += `,\n`;
                    outputContent += `\t\t\t"pre_post": true,\n`;
                    outputContent += `\t\t\t"pre_label": "Out",\n`;
                    outputContent += `\t\t\t"post_label": "In"\n`;
                  }
                  else {
                    outputContent += `\n`;
                  }

                  labelAfter = 0;
                }
              }
            }
            outputContent += `\t\t},\n`;
            id++;
          });
  
          outputContent += `\t],\n`;
          outputContent += `\t"tab_bar_item": {\n`;
          outputContent += `\t\t"id": "",\n`;
          outputContent += `\t\t"url": "",\n`;
          outputContent += `\t\t"title": "${sheetName}",\n`;
          outputContent += `\t\t"image_name": "formDisabled.png"\n`;
          outputContent += `\t}\n`;
          outputContent += `},\n`;
  
          outputContent = outputContent.replace(/,\n\t],/g, '\n\t],');
          outputContent = outputContent.replace(/,\n\t\t\t],/g, '\n\t\t\t],');
  
          let lb6Output = [
            `{
              "label": "${sheetName}",
              "value": "RC2, LB1, LB2, F${fid}", E1, M1, N2, C2"
            },`
          ]

          const outputFilePath = `${sheetName}_${'output'}.txt`;
          
          const blob = new Blob([outputContent], { type: 'text/plain' });
  
          const a = document.createElement('a');
          a.href = URL.createObjectURL(blob);
          a.download = outputFilePath;
          a.click();
  
          URL.revokeObjectURL(a.href);
  
          statusMessage.textContent = 'File processed and downloaded successfully.';
  
          allProcessedData.push([sheetName, outputContent]);
          lb6.push([lb6Output]);
        });

        outputContent += [
              `{
                "id": 91,
                "key": "E1",
                "name": "Email Transaction Summary Module",
                "email": {
                  "id": 92,
                  "label": "Press the Send Email button to send a transaction summary email to the specified recipient(s).",
                  "button": "Send Email",
                  "visible": true,
                  "include_image_inline": true,
                  "send_to_text_default": "{USER}",
                  "include_image_download_link": true,
                  "include_video_download_link": true
                },
                "tab_bar_item": {
                  "id": 93,
                  "url": "",
                  "title": "Send Email",
                  "image_name": "emailDisabled.png"
                }
              },
              {
                "id": 94,
                "key": "E2",
                "name": "Email Transaction Summary Module",
                "email": {
                  "id": 95,
                  "label": "Press the Send Email button to send a transaction summary email to the specified recipient(s).",
                  "button": "Send Email",
                  "visible": true,
                  "include_image_inline": true,
                  "send_to_text_default": "{USER}",
                  "include_image_download_link": true,
                  "include_video_download_link": true
                },
                "tab_bar_item": {
                  "id": 96,
                  "url": "",
                  "title": "Send Email",
                  "image_name": "emailDisabled.png"
                }
              },
              {
                "id": 97,
                "key": "C2",
                "name": "Complete Module 2",
                "label": "Client hereby acknowledges that the attached images/video represent the Condition of Property associated with this transaction at the time of exchange indicated above.",
                "tab_bar_item": {
                  "id": 98,
                  "url": "",
                  "title": "Upload",
                  "image_name": "upload.png"
                }
              },
              {
                "id": 99,
                "key": "LB6",
                "name": "Dynamic Workflow",
                "logic_branch": {
                  "id": 100,
                  "options": [`
        ];

        outputContent += lb6;

        outputContent += [          
                  `],
                  "variable": "{WORKFLOW_MODULES}"
                }
              }
            ],
            "force_upload_dialog_visibility": true
          }
        }`
        ];
  
        allProcessedData.push([outputContent]);

        const combinedOutputFilePath = `${'combined'}.txt`;
  
        const combinedOutputContent = allProcessedData.map(([_, sheetContent]) => sheetContent).join('');
  
        const combinedBlob = new Blob([combinedOutputContent], { type: 'text/plain' });
  
        const aCombined = document.createElement('a');
        aCombined.href = URL.createObjectURL(combinedBlob);
        aCombined.download = combinedOutputFilePath;
        aCombined.click();
  
        URL.revokeObjectURL(aCombined.href);
      };
  
      reader.onerror = (event) => {
        statusMessage.textContent = 'An error occurred while reading the file.';
      };
  
      reader.readAsArrayBuffer(selectedFile);
    });
  });
  