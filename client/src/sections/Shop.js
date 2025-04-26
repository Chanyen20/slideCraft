import React, { useState } from "react";
import styled from "styled-components";

import img15 from "../assets/Images/15.webp";
import img16 from "../assets/Images/16.webp";
import img17 from "../assets/Images/17.webp";
import img18 from "../assets/Images/18.webp";
import img19 from "../assets/Images/19.webp";
import img20 from "../assets/Images/20.webp";

const Section = styled.section`
  min-height: 100vh;
  width: 100vw;
  margin: 0 auto;
  overflow: hidden;
  display: flex;
  justify-content: flex-start;
  align-items: flex-start;
  position: relative;
`;

const Title = styled.h1`
  font-size: ${(props) => props.theme.fontxxxl};
  font-family: "Kaushan Script";
  font-weight: 300;
  text-shadow: 1px 1px 1px ${(props) => props.theme.body};
  color: ${(props) => props.theme.text};
  position: absolute;
  top: 1rem;
  left: 5%;
  z-index: 11;
`;

const GuideText = styled.p`
  font-size: 1.5rem; 
  font-weight: bold; 
  text-align: left;
  line-height: 1.8; 
`;

const Left = styled.div`
  width: 35%;
  background-color: ${(props) => props.theme.body};
  color: ${(props) => props.theme.text};
  min-height: 100vh;
  z-index: 5;
  position: fixed;
  left: 0;
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  padding: 2rem;
`;

// const Right = styled.div`
//   position: absolute;
//   left: 35%;
//   padding-left: 10%;
//   min-height: 100vh;
//   background-color: ${(props) => props.theme.grey};
//   display: flex;
//   flex-direction: column;
//   justify-content: center;
//   align-items: center;
// `;

// const UploadContainer = styled.div`
//   background: white;
//   padding: 2rem;
//   border-radius: 10px;
//   box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.1);
//   text-align: center;
//   width: 80%;
// `;

const Right = styled.div`
  position: absolute;
  left: 35%;
  padding-left: 10%;
  min-height: 100vh;
  background-color: ${(props) => props.theme.grey};
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  gap: 2rem; /* 讓上傳區塊和圖片區塊有間距 */
`;

const UploadContainer = styled.div` 
  background: white;
  padding: 2rem;
  border-radius: 10px;
  box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.1);
  text-align: center;
  width: 80%;
`;

/* 圖片區塊，讓它可以橫向滾動 */
const ImageContainer = styled.div`
  display: flex;
  overflow-x: auto; /* 讓圖片可以橫向滾動 */
  white-space: nowrap;
  width: 100%;
  padding: 1rem;
  gap: 10px; /* 圖片間距 */

  &::-webkit-scrollbar {
    height: 8px;
  }
  &::-webkit-scrollbar-thumb {
    background: ${(props) => props.theme.text};
    border-radius: 4px;
  }
`;

const ImageItem = styled.img`
  width: 150px; /* 控制圖片大小 */
  height: auto;
  border-radius: 10px;
  cursor: pointer;
  transition: transform 0.3s ease;

  &:hover {
    transform: scale(1.1);
  }
`;

const Input = styled.input`
  display: none;
`;

const Label = styled.label`
  background-color: ${(props) => props.theme.text};
  color: ${(props) => props.theme.body};
  padding: 0.8rem 2rem;
  border-radius: 5px;
  cursor: pointer;
  display: inline-block;
  margin-top: 1rem;
`;

const Button = styled.button`
  background-color: ${(props) => props.theme.text};
  color: ${(props) => props.theme.body};
  padding: 0.8rem 2rem;
  border-radius: 5px;
  cursor: pointer;
  margin-top: 1rem;
  border: none;
  font-size: 1rem;
  transition: all 0.3s ease;

  &:hover {
    background-color: ${(props) => props.theme.textLight};
  }
`;

const Shop = () => {
  const [file, setFile] = useState(null);
  const [includeImages, setIncludeImages] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [downloadLink, setDownloadLink] = useState("");
  const [background, setBackground] = useState("");

  const handleFileChange = (e) => {
    setFile(e.target.files[0]);
  };

  const handleUpload = async () => {
    if (!file) {
      alert("Please choose a Word file");
      return;
    }

    setUploading(true);
    setDownloadLink("");

    const formData = new FormData();
    formData.append("file", file);
    formData.append("parse_images", includeImages);
    formData.append("theme", JSON.stringify({
      background: "#FF5733", // 這可以保留或改成你選的顏色
      text: "#000000",
      backgroundImage: background  // 新增這行
    }));

    try {
      const response = await fetch("http://localhost:8000/upload", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error("Fail to upload it!");
      }

      const data = await response.json();
      setDownloadLink(data.pptx_url);
    } catch (error) {
      console.error(error);
      alert("上傳失敗，請稍後再試！");
    } finally {
      setUploading(false);
    }
  };

  return (
    <Section id="mainContent">
      <Title>Upload Your Document</Title>
      <Left>
        <GuideText> 
          Welcome to SlideCraft!  
          <br />
          <br />

            Turn your Word documents into professional presentations effortlessly 
          <br />
          <br />
            How to use?
          <br />
            Step1: Click the button on the right to upload a `.docx` file  
          <br />
            Step2: Choose whether to <strong>include images</strong>  
          <br />
            Step3: Select a <strong>presentation theme  </strong>
          <br />
            Step4: Click <strong>"Generate Slides"</strong>, and let AI handle the rest! 
        </GuideText>
      </Left>
      <Right>
        <UploadContainer>
          <h2>Upload Your Word File</h2>
          <Input type="file" id="fileUpload" accept=".docx" onChange={handleFileChange} />
          <Label htmlFor="fileUpload">Choose File</Label>
          {file && <p> {file.name}</p>}

          <div>
            <label>
              <input
                type="checkbox"
                checked={includeImages}
                onChange={(e) => setIncludeImages(e.target.checked)}
              />{" "}
              Include Images?
            </label>
          </div>

          <Button onClick={handleUpload} disabled={uploading}>
            {uploading ? "Uploading..." : "Generate Slides"}
          </Button>

          {downloadLink && (
            <Button as="a" href={downloadLink} download>
              Download Presentation
            </Button>
          )}
        </UploadContainer>

        <ImageContainer>
          <ImageItem src={img15} alt="Background 1" onClick={() => setBackground("15.webp")} />
          <ImageItem src={img16} alt="Background 2" onClick={() => setBackground("16.webp")} />
          <ImageItem src={img17} alt="Background 3" onClick={() => setBackground("17.webp")} />
          <ImageItem src={img18} alt="Background 3" onClick={() => setBackground("18.webp")} />
          <ImageItem src={img19} alt="Background 3" onClick={() => setBackground("19.webp")} />
          <ImageItem src={img20} alt="Background 3" onClick={() => setBackground("20.webp")} />
        </ImageContainer>
      </Right>
    </Section>
  );
};

export default Shop;
