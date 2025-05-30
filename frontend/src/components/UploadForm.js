import React, { useState } from 'react';
import axios from 'axios';

function UploadForm() {
  const [excels, setExcels] = useState({
    combined_sources: null,
    official_instagram: null,
    official_facebook: null,
    keywords: null
  });
  const [date, setDate] = useState("");
  const [companyName, setCompanyName] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleExcelChange = (file, type) => {
    if (file) {
      // Ensure the file has the correct name and extension
      const newFile = new File([file], `${type}.xlsx`, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      setExcels(prev => ({
        ...prev,
        [type]: newFile
      }));
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);

    // Validate required files
    const missingFiles = Object.entries(excels)
      .filter(([_, file]) => !file)
      .map(([type]) => type);

    if (missingFiles.length > 0) {
      setError(`Please upload all required Excel files: ${missingFiles.join(', ')}`);
      setLoading(false);
      return;
    }

    if (!date) {
      setError('Please select a date');
      setLoading(false);
      return;
    }

    if (!companyName) {
      setError('Please enter a company name');
      setLoading(false);
      return;
    }

    try {
      const formData = new FormData();
      
      // Add Excel files
      Object.entries(excels).forEach(([type, file]) => {
        if (file) {
          formData.append("excel_files", file);
        }
      });

      // Add other form data
      // formData.append("image", image);
      // formData.append("custom_text", customText);
      formData.append("date", date);
      formData.append("company_name", companyName);

      const response = await axios.post("http://localhost:8000/generate-ppt/", formData, {
        responseType: 'blob',
        headers: {
          'Content-Type': 'multipart/form-data',
        },
        validateStatus: function (status) {
          return status < 500; // Accept all status codes less than 500
        }
      });

      if (response.status !== 200) {
        // Handle error response
        const reader = new FileReader();
        reader.onload = () => {
          try {
            const errorData = JSON.parse(reader.result);
            setError(errorData.detail || 'An error occurred while generating the presentation');
          } catch (e) {
            setError('An error occurred while generating the presentation');
          }
        };
        reader.readAsText(response.data);
        return;
      }

      // Download file
      const blob = new Blob([response.data], { 
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" 
      });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = "report.pptx";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Error:', err);
      if (err.response?.data) {
        const reader = new FileReader();
        reader.onload = () => {
          try {
            const errorData = JSON.parse(reader.result);
            setError(errorData.detail || 'An error occurred while generating the presentation');
          } catch (e) {
            setError('An error occurred while generating the presentation');
          }
        };
        reader.readAsText(err.response.data);
      } else {
        setError(err.message || 'An error occurred while generating the presentation');
      }
    } finally {
      setLoading(false);
    }
  };

  return (
    <form onSubmit={handleSubmit}>
      {error && <div style={{ color: 'red', marginBottom: '10px', padding: '10px', border: '1px solid red', borderRadius: '4px' }}>{error}</div>}
      
      <div style={{ marginBottom: '15px' }}>
        <label style={{ display: 'block', marginBottom: '5px' }}>Combined Sources Excel:</label>
        <input 
          type="file" 
          onChange={e => handleExcelChange(e.target.files[0], 'combined_sources')} 
          accept=".xlsx" 
          required 
        />
      </div>
      <div style={{ marginBottom: '15px' }}>
        <label style={{ display: 'block', marginBottom: '5px' }}>Official Instagram Excel:</label>
        <input 
          type="file" 
          onChange={e => handleExcelChange(e.target.files[0], 'official_instagram')} 
          accept=".xlsx" 
          required 
        />
      </div>
      <div style={{ marginBottom: '15px' }}>
        <label style={{ display: 'block', marginBottom: '5px' }}>Official Facebook Excel:</label>
        <input 
          type="file" 
          onChange={e => handleExcelChange(e.target.files[0], 'official_facebook')} 
          accept=".xlsx" 
          required 
        />
      </div>
      <div style={{ marginBottom: '15px' }}>
        <label style={{ display: 'block', marginBottom: '5px' }}>Keywords Excel:</label>
        <input 
          type="file" 
          onChange={e => handleExcelChange(e.target.files[0], 'keywords')} 
          accept=".xlsx" 
          required 
        />
      </div>
      <div style={{ marginBottom: '15px' }}>
        <label style={{ display: 'block', marginBottom: '5px' }}>Date:</label>
        <input 
          type="date" 
          value={date} 
          onChange={e => setDate(e.target.value)} 
          required 
          style={{ padding: '5px' }}
        />
      </div>
      <div style={{ marginBottom: '15px' }}>
        <label style={{ display: 'block', marginBottom: '5px' }}>Company Name:</label>
        <input 
          type="text" 
          value={companyName} 
          onChange={e => setCompanyName(e.target.value)} 
          placeholder="Enter company name" 
          required 
          style={{ padding: '5px', width: '100%' }}
        />
      </div>
      <button 
        type="submit" 
        disabled={loading}
        style={{
          padding: '10px 20px',
          backgroundColor: loading ? '#ccc' : '#007bff',
          color: 'white',
          border: 'none',
          borderRadius: '4px',
          cursor: loading ? 'not-allowed' : 'pointer'
        }}
      >
        {loading ? "Generating..." : "Generate PPT"}
      </button>
    </form>
  );
}

export default UploadForm;
