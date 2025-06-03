import React, { useState } from "react";
import axios from "axios";

function UploadForm() {
  const [excels, setExcels] = useState({
    combined_sources: null,
    official_instagram: null,
    official_facebook: null,
    keywords: null,
  });
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [companyName, setCompanyName] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [positiveLinks, setPositiveLinks] = useState([""]);
  const [negativeLinks, setNegativeLinks] = useState([""]);
  const [positivePosts, setPositivePosts] = useState([{ image: null, link: "" }]);
  const [negativePosts, setNegativePosts] = useState([{ image: null, link: "" }]);
  const [logos, setLogos] = useState({
    company: null,
    mediaEye: null,
    neuroTime: null,
    competitors: []
  });

  const handleExcelChange = (file, type) => {
    if (file) {
      const newFile = new File([file], `${type}.xlsx`, {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      setExcels((prev) => ({
        ...prev,
        [type]: newFile,
      }));
    }
  };

  const handleLogoChange = (file, type, index = null) => {
    if (file) {
      if (type === 'competitors') {
        const newCompetitors = [...logos.competitors];
        newCompetitors[index] = file;
        setLogos(prev => ({
          ...prev,
          competitors: newCompetitors
        }));
      } else {
        setLogos(prev => ({
          ...prev,
          [type]: file
        }));
      }
    }
  };

  const addCompetitorLogo = () => {
    if (logos.competitors.length < 25) {
      setLogos(prev => ({
        ...prev,
        competitors: [...prev.competitors, null]
      }));
    }
  };

  const removeCompetitorLogo = (index) => {
    setLogos(prev => ({
      ...prev,
      competitors: prev.competitors.filter((_, i) => i !== index)
    }));
  };

  const handleLinkChange = (index, value, type) => {
    if (type === "positive") {
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
    if (type === "positive") {
      setPositiveLinks([...positiveLinks, ""]);
    } else {
      setNegativeLinks([...negativeLinks, ""]);
    }
  };

  const handlePostChange = (index, field, value, type) => {
    if (type === "positive") {
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
    if (type === "positive") {
      setPositivePosts([...positivePosts, { image: null, link: "" }]);
    } else {
      setNegativePosts([...negativePosts, { image: null, link: "" }]);
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);

    const missingFiles = Object.entries(excels)
      .filter(([_, file]) => !file)
      .map(([type]) => type);

    if (missingFiles.length > 0) {
      setError(
        `Please upload all required Excel files: ${missingFiles.join(", ")}`
      );
      setLoading(false);
      return;
    }

    if (!startDate || !endDate) {
      setError("Please select both start and end dates");
      setLoading(false);
      return;
    }

    if (!companyName) {
      setError("Please enter a company name");
      setLoading(false);
      return;
    }

    if (!logos.company || !logos.mediaEye || !logos.neuroTime) {
      setError("Please upload all required logos");
      setLoading(false);
      return;
    }

    try {
      const formData = new FormData();

      Object.entries(excels).forEach(([type, file]) => {
        if (file) {
          formData.append("excel_files", file);
        }
      });

      formData.append("company_logo", logos.company);
      formData.append("mediaeye_logo", logos.mediaEye);
      formData.append("neurotime_logo", logos.neuroTime);
      
      // Append competitor logos as an array
      logos.competitors.forEach(logo => {
        if (logo) {
          formData.append("competitor_logos", logo);
        }
      });

      formData.append(
        "positive_links",
        JSON.stringify(positiveLinks.filter((link) => link.trim() !== ""))
      );
      formData.append(
        "negative_links",
        JSON.stringify(negativeLinks.filter((link) => link.trim() !== ""))
      );

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

      formData.append("start_date", startDate);
      formData.append("end_date", endDate);
      formData.append("company_name", companyName);

      const response = await axios.post(
        "http://localhost:8000/generate-ppt/",
        formData,
        {
          responseType: "blob",
          headers: {
            "Content-Type": "multipart/form-data",
          },
          validateStatus: function (status) {
            return status < 500;
          },
        }
      );

      if (response.status !== 200) {
        const reader = new FileReader();
        reader.onload = () => {
          try {
            const errorData = JSON.parse(reader.result);
            setError(errorData.detail || "An error occurred.");
          } catch {
            setError("An error occurred.");
          }
        };
        reader.readAsText(response.data);
        setLoading(false);
        return;
      }

      const blob = new Blob([response.data], {
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "report.pptx";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Error:", err);
      if (err.response?.data) {
        const reader = new FileReader();
        reader.onload = () => {
          try {
            const errorData = JSON.parse(reader.result);
            setError(errorData.detail || "An error occurred.");
          } catch {
            setError("An error occurred.");
          }
        };
        reader.readAsText(err.response.data);
      } else {
        setError(err.message || "An error occurred.");
      }
    } finally {
      setLoading(false);
    }
  };

  return (
    <>
      <style>{`
        * {
          box-sizing: border-box;
        }
        form {
          max-width: 720px;
          margin: 30px auto;
          background-color: #f9fafb;
          padding: 30px 40px;
          border-radius: 8px;
          box-shadow: 0 4px 12px rgb(0 0 0 / 0.1);
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          color: #333;
        }
        h2 {
          margin-bottom: 25px;
          font-weight: 700;
          font-size: 1.8rem;
          text-align: center;
          color: #222;
        }
        label {
          display: block;
          font-weight: 600;
          margin-bottom: 8px;
          color: #444;
        }
        input[type="file"],
        input[type="url"],
        input[type="text"],
        input[type="date"] {
          width: 100%;
          padding: 10px 12px;
          border: 1.5px solid #ccc;
          border-radius: 6px;
          font-size: 1rem;
          transition: border-color 0.3s ease;
        }
        input[type="file"]:focus,
        input[type="url"]:focus,
        input[type="text"]:focus,
        input[type="date"]:focus {
          outline: none;
          border-color: #4a90e2;
          box-shadow: 0 0 6px rgba(74, 144, 226, 0.3);
        }
        .section {
          margin-bottom: 30px;
        }
        .link-list,
        .post-list {
          display: flex;
          flex-direction: column;
          gap: 10px;
        }
        .post-item {
          padding: 15px 20px;
          border: 1px solid #d1d5db;
          border-radius: 8px;
          background-color: #fff;
          box-shadow: 0 2px 6px rgb(0 0 0 / 0.05);
          display: flex;
          flex-direction: column;
        }
        .post-item input[type="file"] {
          margin-bottom: 10px;
        }
        button.add-btn {
          margin-top: 10px;
          background-color: #4a90e2;
          color: white;
          padding: 10px 18px;
          border: none;
          border-radius: 6px;
          font-weight: 600;
          cursor: pointer;
          transition: background-color 0.3s ease;
          width: fit-content;
        }
        button.add-btn:hover:not(:disabled) {
          background-color: #357abd;
        }
        button.add-btn:disabled {
          background-color: #a0b9e4;
          cursor: not-allowed;
        }
        .error-message {
          background-color: #fef2f2;
          color: #b91c1c;
          padding: 12px 18px;
          margin-bottom: 25px;
          border-radius: 6px;
          border: 1px solid #fca5a5;
          font-weight: 600;
          text-align: center;
        }
        button.submit-btn {
          width: 100%;
          background-color: #2563eb;
          color: white;
          padding: 15px 0;
          font-size: 1.1rem;
          font-weight: 700;
          border: none;
          border-radius: 8px;
          cursor: pointer;
          transition: background-color 0.3s ease;
        }
        button.submit-btn:hover:not(:disabled) {
          background-color: #1e40af;
        }
        button.submit-btn:disabled {
          background-color: #93c5fd;
          cursor: not-allowed;
        }
        @media (max-width: 600px) {
          form {
            padding: 20px 25px;
          }
        }
      `}</style>

      <form onSubmit={handleSubmit} noValidate>
        <h2>Upload Required Files & Data</h2>

        {error && <div className="error-message">{error}</div>}

        {/* Excel Upload Section */}
        <div
          className="section"
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            gap: "20px",
            marginTop: "10px"
          }}
        >
          <div>
            <label htmlFor="combined_sources">Combined Sources</label>
            <input
              id="combined_sources"
              type="file"
              accept=".xlsx"
              onChange={(e) => handleExcelChange(e.target.files?.[0], "combined_sources")}
              aria-label="Combined sources Excel file"
              style={{ marginTop: "5px" }}
            />
          </div>

          <div>
            <label htmlFor="official_instagram">Official Instagram</label>
            <input
              id="official_instagram"
              type="file"
              accept=".xlsx"
              onChange={(e) => handleExcelChange(e.target.files?.[0], "official_instagram")}
              aria-label="Official Instagram Excel file"
              style={{ marginTop: "5px" }}
            />
          </div>

          <div>
            <label htmlFor="official_facebook">Official Facebook</label>
            <input
              id="official_facebook"
              type="file"
              accept=".xlsx"
              onChange={(e) => handleExcelChange(e.target.files?.[0], "official_facebook")}
              aria-label="Official Facebook Excel file"
              style={{ marginTop: "5px" }}
            />
          </div>

          <div>
            <label htmlFor="keywords">Keywords</label>
            <input
              id="keywords"
              type="file"
              accept=".xlsx"
              onChange={(e) => handleExcelChange(e.target.files?.[0], "keywords")}
              aria-label="Keywords Excel file"
              style={{ marginTop: "5px" }}
            />
          </div>
        </div>

        {/* Logos Section */}
        <div className="section">
          <h3>Logos</h3>
          <div className="logo-inputs">
            <div>
              <label>Company Logo:</label>
              <input
                type="file"
                accept="image/*"
                onChange={(e) => handleLogoChange(e.target.files[0], 'company')}
                required
              />
            </div>
            <div>
              <label>MediaEye Logo:</label>
              <input
                type="file"
                accept="image/*"
                onChange={(e) => handleLogoChange(e.target.files[0], 'mediaEye')}
                required
              />
            </div>
            <div>
              <label>NeuroTime Logo:</label>
              <input
                type="file"
                accept="image/*"
                onChange={(e) => handleLogoChange(e.target.files[0], 'neuroTime')}
                required
              />
            </div>
          </div>

          {/* Competitor Logos */}
          <div className="competitor-logos">
            <label>Competitor Logos (up to 25):</label>
            {logos.competitors.map((logo, index) => (
              <div key={index} className="competitor-logo-item">
                <input
                  type="file"
                  accept="image/*"
                  onChange={(e) => handleLogoChange(e.target.files[0], 'competitors', index)}
                />
                <button
                  type="button"
                  onClick={() => removeCompetitorLogo(index)}
                  className="remove-btn"
                >
                  Remove
                </button>
              </div>
            ))}
            {logos.competitors.length < 25 && (
              <button
                type="button"
                onClick={addCompetitorLogo}
                className="add-btn"
              >
                + Add Competitor Logo
              </button>
            )}
          </div>
        </div>

        {/* Date Range Section */}
        <div className="section">
          <h3>Date Range</h3>
          <div className="date-inputs">
            <div>
              <label>Start Date:</label>
              <input
                type="date"
                value={startDate}
                onChange={(e) => setStartDate(e.target.value)}
                required
              />
            </div>
            <div>
              <label>End Date:</label>
              <input
                type="date"
                value={endDate}
                onChange={(e) => setEndDate(e.target.value)}
                required
              />
            </div>
          </div>
        </div>

        {/* Company Name Section */}
        <div className="section">
          <label>Company Name:</label>
          <input
            type="text"
            id="company"
            placeholder="Enter company name"
            value={companyName}
            onChange={(e) => setCompanyName(e.target.value)}
            required
          />
        </div>

        {/* Positive Links Section */}
        <div className="section">
          <label>Positive Links:</label>
          <div className="link-list">
            {positiveLinks.map((link, i) => (
              <input
                key={`positive-link-${i}`}
                type="url"
                placeholder={`Positive Link #${i + 1}`}
                value={link}
                onChange={(e) => handleLinkChange(i, e.target.value, "positive")}
              />
            ))}
          </div>
          <button
            type="button"
            className="add-btn"
            onClick={() => addLinkField("positive")}
          >
            + Add Positive Link
          </button>
        </div>

        {/* Negative Links Section */}
        <div className="section">
          <label>Negative Links:</label>
          <div className="link-list">
            {negativeLinks.map((link, i) => (
              <input
                key={`negative-link-${i}`}
                type="url"
                placeholder={`Negative Link #${i + 1}`}
                value={link}
                onChange={(e) => handleLinkChange(i, e.target.value, "negative")}
              />
            ))}
          </div>
          <button
            type="button"
            className="add-btn"
            onClick={() => addLinkField("negative")}
          >
            + Add Negative Link
          </button>
        </div>

        {/* Positive Posts Section */}
        <div className="section">
          <label>Positive Posts (Image + Link):</label>
          <div className="post-list">
            {positivePosts.map((post, i) => (
              <div key={`positive-post-${i}`} className="post-item">
                <input
                  type="file"
                  accept="image/*"
                  onChange={(e) =>
                    handlePostChange(i, "image", e.target.files?.[0], "positive")
                  }
                  aria-label={`Positive post image #${i + 1}`}
                />
                <input
                  type="url"
                  placeholder="Post Link"
                  value={post.link}
                  onChange={(e) =>
                    handlePostChange(i, "link", e.target.value, "positive")
                  }
                />
              </div>
            ))}
          </div>
          <button
            type="button"
            className="add-btn"
            onClick={() => addPostField("positive")}
          >
            + Add Positive Post
          </button>
        </div>

        {/* Negative Posts Section */}
        <div className="section">
          <label>Negative Posts (Image + Link):</label>
          <div className="post-list">
            {negativePosts.map((post, i) => (
              <div key={`negative-post-${i}`} className="post-item">
                <input
                  type="file"
                  accept="image/*"
                  onChange={(e) =>
                    handlePostChange(i, "image", e.target.files?.[0], "negative")
                  }
                  aria-label={`Negative post image #${i + 1}`}
                />
                <input
                  type="url"
                  placeholder="Post Link"
                  value={post.link}
                  onChange={(e) =>
                    handlePostChange(i, "link", e.target.value, "negative")
                  }
                />
              </div>
            ))}
          </div>
          <button
            type="button"
            className="add-btn"
            onClick={() => addPostField("negative")}
          >
            + Add Negative Post
          </button>
        </div>

        <button
          type="submit"
          className="submit-btn"
          disabled={loading}
          aria-busy={loading}
        >
          {loading ? "Generating..." : "Generate Report"}
        </button>
      </form>
    </>
  );
}

export default UploadForm;
