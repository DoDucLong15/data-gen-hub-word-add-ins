// Định nghĩa kiểu dữ liệu cho file đặc tả
interface MappingCell {
  cell: string;
  dbfield?: string;
  dbfields?: string[];
  const?: string;
  comment?: string;
  identity?: number;
}

interface Mapping {
  dbtablename: string;
  cells: MappingCell[];
  _comment_mapping_?: string;
  _comment_dbtablename_?: string;
}

interface Config {
  nameformat: string[];
  _comment_nameformat?: string;
}

interface SpecFile {
  config: Config;
  document: {
    mapping: Mapping;
  };
  errMessage: string;
}

// Khai báo biến toàn cục
let spec: SpecFile = {
  config: {
    nameformat: [],
    _comment_nameformat:
      "Kí tự đầu tiên ? báo hiệu lấy theo dbfiled, nếu không có là hằng kí tự. Chỉ áp dụng cho xuất file từ db-->excel",
  },
  document: {
    mapping: {
      dbtablename: "",
      _comment_dbtablename_:
        "Tên bảng dữ liệu trong cơ sở dữ liệu. Riêng đối với db2excel thì các Sheet phải cùng lấy só liệu từ 1 bảng dữ liệu và là của sheet đầu tiên",
      cells: [],
      _comment_mapping_:
        'dbfield mô tả trường dữ liệu đơn. dbfields mô tả trường dữ liệu phức, ghép xâu các trường đơn để tổng hợp dữ liệu, ví dụ dbfields:["{0}...{1}...{2}","gvhd","tendoan","loaidoan"] ',
    },
  },
  errMessage: "",
};

// Add this interface and predefined data near the top of your TypeScript file
interface TableFields {
  [tableName: string]: string[]; // Tên bảng -> danh sách các trường
}

// Khai báo dữ liệu mẫu
const predefinedTables: TableFields = {
  students: [
    "student_class_name",
    "supervisor",
    "reviewer",
    "phone",
    "class_id",
    "mssv",
    "updated_at",
    "middle_name",
    "first_name",
    "email",
    "last_name",
    "project_title",
  ],
  assignment_sheets: [
    "semester",
    "class_code",
    "expected_products",
    "thesis_start_date",
    "student_sign_date",
    "email",
    "input_path",
    "real_world_problem_solved",
    "student_class_name",
    "thesis_end_date",
    "phone",
    "school",
    "mssv",
    "supervisor_sign_date",
    "project_title",
    "technology_gained",
    "full_name",
    "supervisor",
    "student_knowledge_gained",
    "acquired_skills",
  ],
  guidance_reviews: [
    "problem_difficulty_point",
    "type_of_thesis",
    "response_accuracy_point",
    "topic_uniqueness_point",
    "product_finalization_point",
    "literature_review_point",
    "workload_point",
    "reward_point",
    "layout_coherence_point",
    "mssv",
    "presentation_skills_point",
    "general_feedback",
    "solution_impact_point",
    "project_title",
    "full_name",
    "supervisor",
    "teacher_sign_date",
    "conclusion",
    "presentation_quality_point",
    "content_validity_point",
  ],
  supervisory_comments: ["supervisor", "full_name", "mssv", "project_title"],
};

// Khởi tạo khi add-in được tải
// Đăng ký sự kiện cho Word
Office.onReady(() => {
  // Populate table dropdown
  populateTableDropdown();

  // Initial update of field dropdowns
  updateFieldOptions();
  updateDbFieldsOptions();

  // Add event listener for dbfieldsSelect
  const dbfieldsSelect = document.getElementById("dbfieldsSelect") as HTMLSelectElement;
  if (dbfieldsSelect) {
    dbfieldsSelect.addEventListener("change", updateSelectedDbFields);
  }

  // Cập nhật lần đầu
  updateSelectedRange().catch(console.error);

  Word.run(async (context: Word.RequestContext) => {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
      updateSelectedRange().catch(console.error);
    });

    await context.sync();
  }).catch(console.error);

  updateSpecDisplay(); // Hiển thị dữ liệu ban đầu
});

// Xử lý khi selection thay đổi
async function handleSelectionChange(): Promise<void> {
  await updateSelectedRange();
}

// Cập nhật vùng đang chọn
async function updateSelectedRange(): Promise<void> {
  try {
    await Word.run(async (context: Word.RequestContext) => {
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();

      // Hiển thị một phần của text được chọn (nếu quá dài)
      let displayText = range.text.trim();
      if (displayText.length > 30) {
        displayText = displayText.substring(0, 30) + "...";
      }
      const selectionInfo = displayText ? `"${displayText}"` : "None";

      document.querySelectorAll(".selected-range").forEach((el: Element) => {
        (el as HTMLElement).textContent = selectionInfo;
      });
    });
  } catch (error) {
    console.error("Error updating selected range:", error);
  }
}

// Cập nhật hiển thị dữ liệu spec
function updateSpecDisplay(): void {
  const specData: HTMLElement | null = document.getElementById("spec-data");
  if (specData) {
    // Format JSON với indent 2 spaces và thêm màu sắc
    const formattedJson = JSON.stringify(spec, null, 2)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(
        /("(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\"])*"(\s*:)?|\b(true|false|null)\b|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?)/g,
        function (match) {
          let cls = "number";
          if (/^"/.test(match)) {
            if (/:$/.test(match)) {
              cls = "key";
            } else {
              cls = "string";
            }
          } else if (/true|false/.test(match)) {
            cls = "boolean";
          } else if (/null/.test(match)) {
            cls = "null";
          }
          return '<span class="' + cls + '">' + match + "</span>";
        }
      );

    specData.innerHTML = formattedJson;
  }

  const nameFormatDisplay: HTMLElement | null = document.getElementById("nameFormatDisplay");
  if (nameFormatDisplay) {
    nameFormatDisplay.textContent =
      spec.config.nameformat.length > 0 ? spec.config.nameformat.join("") : "Chưa có định dạng";
  }
}

// Thêm nameFormat vào config
function addNameFormat(): void {
  const nameFormatInput: HTMLInputElement | null = document.getElementById(
    "nameFormat"
  ) as HTMLInputElement;
  const nameFormat: string = nameFormatInput?.value.trim() ?? "";
  if (nameFormat) {
    spec.config.nameformat = nameFormat.split(/(?=\?)/);
    console.log("Config updated:", spec.config);
    updateSpecDisplay();
  } else {
    alert("Vui lòng nhập định dạng tên!");
  }
}

// Hiển thị input tương ứng với mapping type
function toggleMappingInputs(): void {
  const mappingTypeSelect: HTMLSelectElement | null = document.getElementById(
    "mappingType"
  ) as HTMLSelectElement;
  const mappingType: string = mappingTypeSelect?.value ?? "dbfield";
  const inputs: string[] = ["dbfieldInput", "dbfieldsInput", "constInput", "commentInput"];

  inputs.forEach((id: string) => {
    const element: HTMLElement | null = document.getElementById(id);
    if (element) element.style.display = "none";
  });

  const activeInput: HTMLElement | null = document.getElementById(`${mappingType}Input`);
  if (activeInput) activeInput.style.display = "block";

  // Update field options for the active input type
  if (mappingType === "dbfield") {
    updateFieldOptions();
  } else if (mappingType === "dbfields") {
    updateDbFieldsOptions();
  }
}

// Cập nhật danh sách mapping
function updateMappingList(): void {
  const mappingList: HTMLElement | null = document.getElementById("mappingList");
  if (!mappingList) return;

  mappingList.innerHTML = "";
  spec.document.mapping.cells.forEach((mapping: MappingCell, mappingIndex: number) => {
    const item: HTMLDivElement = document.createElement("div");
    item.className = "mapping-item";

    // Đảm bảo hiển thị rõ ràng cell value
    let cellValue = mapping.cell;
    let mappingText: string = `${cellValue}: `;

    if (mapping.dbfield) {
      mappingText += `dbfield=${mapping.dbfield}`;
      if (mapping.identity) {
        mappingText += `, identity=${mapping.identity}`;
      }
    } else if (mapping.dbfields) {
      mappingText += `dbfields=[${mapping.dbfields.join(", ")}]`;
    } else if (mapping.const) {
      mappingText += `const="${mapping.const}"`;
    } else if (mapping.comment) {
      mappingText += `comment="${mapping.comment}"`;
    }

    // Sử dụng textContent thay vì innerHTML cho span để tránh lỗi render HTML
    const textSpan = document.createElement("span");
    textSpan.textContent = mappingText;

    // Tạo button xóa
    const deleteButton = document.createElement("button");
    deleteButton.textContent = "Delete";
    deleteButton.onclick = function () {
      deleteMapping(mappingIndex);
    };

    // Thêm các phần tử vào item
    item.appendChild(textSpan);
    item.appendChild(deleteButton);

    // Thêm item vào danh sách
    mappingList.appendChild(item);
  });
}

// Thêm mapping
async function addMapping(): Promise<void> {
  try {
    await Word.run(async (context: Word.RequestContext) => {
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();

      // Sử dụng text được chọn làm field name
      let fieldName = range.text.trim();
      if (!fieldName) {
        alert("Vui lòng chọn vùng văn bản để thêm mapping!");
        return;
      }

      // Giới hạn độ dài và format field name
      if (fieldName.length > 50) {
        fieldName = fieldName.substring(0, 50) + "...";
      }

      // Đặt trong <field> để phù hợp với yêu cầu
      const cellValue = `<${fieldName}>`;

      // Cập nhật dbTableName
      const dbTableNameInput: HTMLInputElement | null = document.getElementById(
        "dbTableName"
      ) as HTMLInputElement;
      const dbTableName = dbTableNameInput?.value.trim() ?? "";
      spec.document.mapping.dbtablename = dbTableName;

      if (!dbTableName) {
        alert("Vui lòng nhập tên bảng dữ liệu (DB Table Name)!");
        return;
      }

      const mappingTypeSelect: HTMLSelectElement | null = document.getElementById(
        "mappingType"
      ) as HTMLSelectElement;
      const mappingType: string = mappingTypeSelect?.value ?? "dbfield";
      const mapping: MappingCell = { cell: cellValue };

      switch (mappingType) {
        case "dbfield":
          const dbfieldSelect = document.getElementById("dbfield") as HTMLSelectElement;
          const dbfield = dbfieldSelect?.value ?? "";
          const identityCheckbox = document.getElementById("identityCheckbox") as HTMLInputElement;

          if (dbfield) {
            mapping.dbfield = dbfield;
            if (identityCheckbox?.checked) {
              mapping.identity = 1;
            }
          } else {
            alert("Vui lòng chọn tên DB Field!");
            return;
          }
          break;

        case "dbfields":
          const formatInput = document.getElementById("dbfieldsFormat") as HTMLInputElement;
          const fieldsInput = document.getElementById("dbfieldsList") as HTMLInputElement;
          const format = formatInput?.value.trim() ?? "";
          const fieldsValue = fieldsInput?.value.trim() ?? "";

          if (!format) {
            alert("Vui lòng nhập định dạng cho DB Fields!");
            return;
          }

          if (!fieldsValue) {
            alert("Vui lòng chọn các DB Fields!");
            return;
          }

          const fields = fieldsValue.split(",").map((f) => f.trim());
          if (format && fields.length) {
            mapping.dbfields = [format, ...fields];
          }
          break;

        case "const":
          const constInput: HTMLInputElement | null = document.getElementById(
            "constValue"
          ) as HTMLInputElement;
          const constValue: string = constInput?.value.trim() ?? "";

          if (!constValue) {
            alert("Vui lòng nhập giá trị hằng số!");
            return;
          }

          mapping.const = constValue;
          break;

        case "comment":
          const commentInput: HTMLTextAreaElement | null = document.getElementById(
            "commentValue"
          ) as HTMLTextAreaElement;
          const commentValue: string = commentInput?.value.trim() ?? "";

          if (!commentValue) {
            alert("Vui lòng nhập nội dung comment!");
            return;
          }

          mapping.comment = commentValue;
          break;
      }

      // Thêm mapping vào spec
      spec.document.mapping.cells.push(mapping);
      clearMappingInputs();
      console.log("Mapping added:", mapping);
      updateMappingList();
      updateSpecDisplay();
    });
  } catch (error) {
    console.error("Error adding mapping:", error);
    alert(
      "Đã xảy ra lỗi khi thêm mapping: " +
        (error instanceof Error ? error.message : "Unknown error")
    );
  }
}

// Xóa mapping
function deleteMapping(mappingIndex: number): void {
  spec.document.mapping.cells.splice(mappingIndex, 1);
  updateMappingList();
  updateSpecDisplay();
}

// Xóa input sau khi thêm mapping
function clearMappingInputs(): void {
  const inputs: string[] = [
    "dbfield",
    "dbfieldsFormat",
    "dbfieldsList",
    "constValue",
    "commentValue",
  ];

  inputs.forEach((id: string) => {
    const element: HTMLInputElement | HTMLTextAreaElement | null = document.getElementById(id) as
      | HTMLInputElement
      | HTMLTextAreaElement;
    if (element) element.value = "";
  });

  const identityCheckbox: HTMLInputElement | null = document.getElementById(
    "identityCheckbox"
  ) as HTMLInputElement;
  if (identityCheckbox) identityCheckbox.checked = false;
}

// Sinh file đặc tả
function generateSpecFile(): void {
  // Cập nhật dbTableName
  const dbTableNameInput: HTMLInputElement | null = document.getElementById(
    "dbTableName"
  ) as HTMLInputElement;
  spec.document.mapping.dbtablename = dbTableNameInput?.value.trim() ?? "";

  if (spec.document.mapping.cells.length === 0 && spec.config.nameformat.length === 0) {
    alert("Vui lòng thêm dữ liệu để tạo file đặc tả!");
    return;
  }

  const jsonString: string = JSON.stringify(spec, null, 2);
  const blob: Blob = new Blob([jsonString], { type: "application/json" });
  const link: HTMLAnchorElement = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "spec.json";
  link.click();

  // Hiển thị thông báo thành công
  const notification = document.createElement("div");
  notification.className = "success-notification";
  notification.textContent = "File đặc tả đã được tạo thành công!";
  document.body.appendChild(notification);

  setTimeout(() => {
    notification.remove();
  }, 3000);
}

// Lưu tên bảng DB
function saveDbTableName(): void {
  const dbTableNameInput: HTMLInputElement | null = document.getElementById(
    "dbTableName"
  ) as HTMLInputElement;
  spec.document.mapping.dbtablename = dbTableNameInput?.value.trim() ?? "";

  // Update field dropdowns based on the selected table
  updateFieldOptions();
  updateDbFieldsOptions();

  // Hiển thị thông báo
  const notification = document.createElement("div");
  notification.className = "success-notification";
  notification.textContent = "Đã lưu tên bảng DB!";
  document.body.appendChild(notification);

  setTimeout(() => {
    notification.remove();
  }, 3000);

  updateSpecDisplay();
}

// Add these functions to your TypeScript file

// Populate table dropdown on startup
function populateTableDropdown(): void {
  const tableSelect = document.getElementById("dbTableName") as HTMLSelectElement;
  if (!tableSelect) return;

  // Clear existing options except the first one
  while (tableSelect.options.length > 1) {
    tableSelect.remove(1);
  }

  // Add predefined tables
  Object.keys(predefinedTables).forEach((tableName) => {
    const option = document.createElement("option");
    option.value = tableName;
    option.textContent = tableName;
    tableSelect.appendChild(option);
  });
}

// Update field dropdowns based on selected table
function updateFieldOptions(): void {
  const tableSelect = document.getElementById("dbTableName") as HTMLSelectElement;
  const dbfieldSelect = document.getElementById("dbfield") as HTMLSelectElement;

  if (!tableSelect || !dbfieldSelect) return;

  const selectedTable = tableSelect.value;

  // Clear existing field options
  while (dbfieldSelect.options.length > 1) {
    dbfieldSelect.remove(1);
  }

  // If no table selected, return
  if (!selectedTable || !predefinedTables[selectedTable]) return;

  // Add fields for the selected table
  predefinedTables[selectedTable].forEach((field) => {
    const option = document.createElement("option");
    option.value = field;
    option.textContent = field;
    dbfieldSelect.appendChild(option);
  });

  // Update the spec with the selected table name
  spec.document.mapping.dbtablename = selectedTable;
  updateSpecDisplay();
}

function updateDbFieldsOptions(): void {
  const tableSelect = document.getElementById("dbTableName") as HTMLSelectElement;
  const dbfieldsSelect = document.getElementById("dbfieldsSelect") as HTMLSelectElement;

  if (!tableSelect || !dbfieldsSelect) return;

  const selectedTable = tableSelect.value;

  // Clear existing field options
  while (dbfieldsSelect.options.length > 1) {
    dbfieldsSelect.remove(1);
  }

  // If no table selected, return
  if (!selectedTable || !predefinedTables[selectedTable]) return;

  // Add fields for the selected table
  predefinedTables[selectedTable].forEach((field) => {
    const option = document.createElement("option");
    option.value = field;
    option.textContent = field;
    dbfieldsSelect.appendChild(option);
  });
}

// Update dbfieldsList input when fields are selected
function updateSelectedDbFields(): void {
  const dbfieldsSelect = document.getElementById("dbfieldsSelect") as HTMLSelectElement;
  const dbfieldsList = document.getElementById("dbfieldsList") as HTMLInputElement;

  if (!dbfieldsSelect || !dbfieldsList) return;

  const selectedFields: string[] = [];

  // Get all selected options
  for (let i = 0; i < dbfieldsSelect.options.length; i++) {
    if (dbfieldsSelect.options[i].selected && dbfieldsSelect.options[i].value) {
      selectedFields.push(dbfieldsSelect.options[i].value);
    }
  }

  // Update the dbfieldsList input with comma-separated values
  dbfieldsList.value = selectedFields.join(", ");
}

// Gắn các hàm vào window
(window as any).addNameFormat = addNameFormat;
(window as any).toggleMappingInputs = toggleMappingInputs;
(window as any).addMapping = addMapping;
(window as any).deleteMapping = deleteMapping;
(window as any).generateSpecFile = generateSpecFile;
(window as any).saveDbTableName = saveDbTableName;
(window as any).updateFieldOptions = updateFieldOptions;
(window as any).updateDbFieldsOptions = updateDbFieldsOptions;
(window as any).updateSelectedDbFields = updateSelectedDbFields;
