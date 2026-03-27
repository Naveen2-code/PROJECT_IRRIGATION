let currentFile = null;
const FILE_NAME = "project.docx";

// Auto load your DOCX
window.onload = () => {
    fetch(FILE_NAME)
        .then(res => res.arrayBuffer())
        .then(buffer => {
            currentFile = buffer;
            renderDoc(buffer);
        })
        .catch(() => {
            document.getElementById("viewer").innerHTML =
                "<p>No default document found. Upload one.</p>";
        });
};

// Upload new DOCX
document.getElementById("fileInput").addEventListener("change", e => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = evt => {
        currentFile = evt.target.result;
        renderDoc(currentFile);
    };
    reader.readAsArrayBuffer(file);
});

// Render DOCX
function renderDoc(buffer) {
    mammoth.convertToHtml({ arrayBuffer: buffer })
        .then(res => {
            document.getElementById("viewer").innerHTML = res.value;
        })
        .catch(() => {
            document.getElementById("viewer").innerHTML =
                "<p style='color:red'>Error reading file</p>";
        });
}

// Download loaded DOCX
document.getElementById("downloadCurrent").onclick = () => {
    if (!currentFile) return alert("No file loaded");

    const blob = new Blob([currentFile], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });

    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "Irrigation_Report.docx";
    a.click();
};

// Export HTML
document.getElementById("downloadHTML").onclick = () => {
    const content = document.getElementById("viewer").innerHTML;

    const blob = new Blob([`
        <html><body>${content}</body></html>
    `], { type: "text/html" });

    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "Report.html";
    a.click();
};