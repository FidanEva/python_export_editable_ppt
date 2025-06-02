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
  const [positiveLinks, setPositiveLinks] = useState([""]);
  const [negativeLinks, setNegativeLinks] = useState([""]);
  const [positivePosts, setPositivePosts] = useState([{ image: null, link: "" }]);
  const [negativePosts, setNegativePosts] = useState([{ image: null, link: "" }]);

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

  const handleLinkChange = (index, value, type) => {
    if (type === 'positive') {
      const newLinks = [...positiveLinks];
      newLinks[index] = value;
      setPositiveLinks(newLinks);
    } else {
      const newLinks = [...negativeLinks];
      newLinks[index] = value;
      setNegativeLinks(newLinks);
    }
  };

  const addLinkField = (type) => {
    if (type === 'positive') {
      setPositiveLinks([...positiveLinks, ""]);
    } else {
      setNegativeLinks([...negativeLinks, ""]);
    }
  };

  const handlePostChange = (index, field, value, type) => {
    if (type === 'positive') {
      const newPosts = [...positivePosts];
      newPosts[index] = { ...newPosts[index], [field]: value };
      setPositivePosts(newPosts);
    } else {
      const newPosts = [...negativePosts];
      newPosts[index] = { ...newPosts[index], [field]: value };
      setNegativePosts(newPosts);
    }
  };

  const addPostField = (type) => {
    if (type === 'positive') {
      setPositivePosts([...positivePosts, { image: null, link: "" }]);
    } else {
      setNegativePosts([...negativePosts, { image: null, link: "" }]);
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

      // Add links
      formData.append("positive_links", JSON.stringify(positiveLinks.filter(link => link.trim() !== "")));
      formData.append("negative_links", JSON.stringify(negativeLinks.filter(link => link.trim() !== "")));

      // Add posts with images
      positivePosts.forEach((post, index) => {
        if (post.image) {
          formData.append(`positive_post_image_${index}`, post.image);
          formData.append(`positive_post_link_${index}`, post.link);
        }
      });

      negativePosts.forEach((post, index) => {
        if (post.image) {
          formData.append(`negative_post_image_${index}`, post.image);
          formData.append(`negative_post_link_${index}`, post.link);
        }
      });

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
        <label style={{ display: 'block', marginBottom: '5px' }}>Positive Links:</label>
        {positiveLinks.map((link, index) => (
          <div key={index} style={{ marginBottom: '5px' }}>
            <input
              type="url"
              value={link}
              onChange={e => handleLinkChange(index, e.target.value, 'positive')}
              placeholder="Enter positive link"
              style={{ width: '100%', padding: '5px' }}
            />
          </div>
        ))}
        <button type="button" onClick={() => addLinkField('positive')} style={{ marginTop: '5px' }}>
          Add Another Positive Link
        </button>
      </div>

      <div style={{ marginBottom: '15px' }}>
        <label style={{ display: 'block', marginBottom: '5px' }}>Negative Links:</label>
        {negativeLinks.map((link, index) => (
          <div key={index} style={{ marginBottom: '5px' }}>
            <input
              type="url"
              value={link}
              onChange={e => handleLinkChange(index, e.target.value, 'negative')}
              placeholder="Enter negative link"
              style={{ width: '100%', padding: '5px' }}
            />
          </div>
        ))}
        <button type="button" onClick={() => addLinkField('negative')} style={{ marginTop: '5px' }}>
          Add Another Negative Link
        </button>
      </div>

      <div style={{ marginBottom: '15px' }}>
        <label style={{ display: 'block', marginBottom: '5px' }}>Positive Posts:</label>
        {positivePosts.map((post, index) => (
          <div key={index} style={{ marginBottom: '15px', padding: '10px', border: '1px solid #ddd', borderRadius: '4px' }}>
            <input
              type="file"
              onChange={e => handlePostChange(index, 'image', e.target.files[0], 'positive')}
              accept="image/*"
              style={{ marginBottom: '5px' }}
            />
            <input
              type="url"
              value={post.link}
              onChange={e => handlePostChange(index, 'link', e.target.value, 'positive')}
              placeholder="Enter post link"
              style={{ width: '100%', padding: '5px' }}
            />
          </div>
        ))}
        <button type="button" onClick={() => addPostField('positive')} style={{ marginTop: '5px' }}>
          Add Another Positive Post
        </button>
      </div>

      <div style={{ marginBottom: '15px' }}>
        <label style={{ display: 'block', marginBottom: '5px' }}>Negative Posts:</label>
        {negativePosts.map((post, index) => (
          <div key={index} style={{ marginBottom: '15px', padding: '10px', border: '1px solid #ddd', borderRadius: '4px' }}>
            <input
              type="file"
              onChange={e => handlePostChange(index, 'image', e.target.files[0], 'negative')}
              accept="image/*"
              style={{ marginBottom: '5px' }}
            />
            <input
              type="url"
              value={post.link}
              onChange={e => handlePostChange(index, 'link', e.target.value, 'negative')}
              placeholder="Enter post link"
              style={{ width: '100%', padding: '5px' }}
            />
          </div>
        ))}
        <button type="button" onClick={() => addPostField('negative')} style={{ marginTop: '5px' }}>
          Add Another Negative Post
        </button>
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
