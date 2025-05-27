import React, { useState } from 'react';
import axios from 'axios';

function UploadForm() {
  const [excels, setExcels] = useState([]);
  const [image, setImage] = useState(null);
  const [customText, setCustomText] = useState("");

  const handleSubmit = async (e) => {
    e.preventDefault();

    const formData = new FormData();
    for (const file of excels) {
      formData.append("excel_files", file);
    }
    formData.append("image", image);
    formData.append("custom_text", customText);

    const response = await axios.post("http://localhost:8000/generate-ppt/", formData, {
      responseType: 'blob',
    });

    // Download file
    const blob = new Blob([response.data], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = "report.pptx";
    a.click();
  };

  return (
    <form onSubmit={handleSubmit}>
      <input type="file" multiple onChange={e => setExcels([...e.target.files])} accept=".xlsx" />
      <input type="file" onChange={e => setImage(e.target.files[0])} accept="image/*" />
      <input type="text" value={customText} onChange={e => setCustomText(e.target.value)} placeholder="Enter custom text" />
      <button type="submit">Generate PPT</button>
    </form>
  );
}

export default UploadForm;
