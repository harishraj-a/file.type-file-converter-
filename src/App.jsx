import { useState } from 'react'
import { renderAsync } from 'docx-preview'
import html2canvas from 'html2canvas'
import { jsPDF } from 'jspdf'
import * as pdfjsLib from 'pdfjs-dist'
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import { Document, Paragraph, TextRun, Packer, PageBreak } from 'docx'
import './App.css'

// Configure PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker

// Supported file formats and metadata
const SUPPORTED_FORMATS = {
  png: { label: 'PNG Image', category: 'image' },
  jpg: { label: 'JPEG Image', category: 'image' },
  jpeg: { label: 'JPEG Image', category: 'image' },
  webp: { label: 'WebP Image', category: 'image' },
  pdf: { label: 'PDF Document', category: 'document' },
  docx: { label: 'Word Document (.docx)', category: 'document' },
}

// Target format options for images
const IMAGE_CONVERSION_TARGETS = {
  png: ['jpg', 'webp'],
  jpg: ['png', 'webp'],
  jpeg: ['png', 'webp'],
  webp: ['png', 'jpg'],
}

// MIME types for image export
const MIME_TYPES = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  webp: 'image/webp',
}

// Format bytes into readable text
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes'
  const units = ['Bytes', 'KB', 'MB', 'GB']
  const i = Math.floor(Math.log(bytes) / Math.log(1024))
  return `${(bytes / Math.pow(1024, i)).toFixed(2)} ${units[i]}`
}

function App() {
  const [selectedFile, setSelectedFile] = useState(null)
  const [fileDetails, setFileDetails] = useState(null)
  const [errorMessage, setErrorMessage] = useState('')
  const [isDragging, setIsDragging] = useState(false)

  // Conversion states
  const [targetFormat, setTargetFormat] = useState('')
  const [isConverting, setIsConverting] = useState(false)
  const [conversionError, setConversionError] = useState('')
  const [convertedResult, setConvertedResult] = useState(null)

  // Process and validate selected file
  const handleFileProcess = (file) => {
    if (!file) return

    // Clean up previous conversion if any
    if (convertedResult?.url) {
      URL.revokeObjectURL(convertedResult.url)
    }
    setConvertedResult(null)
    setConversionError('')

    const extension = file.name.split('.').pop().toLowerCase()
    const matchedFormat = SUPPORTED_FORMATS[extension]

    if (!matchedFormat) {
      setSelectedFile(null)
      setFileDetails(null)
      setErrorMessage(
        `Unsupported file type: ".${extension}". Please upload PNG, JPG, WEBP, PDF, or DOCX.`
      )
      return
    }

    setErrorMessage('')
    setSelectedFile(file)
    setFileDetails({
      name: file.name,
      size: formatFileSize(file.size),
      typeLabel: matchedFormat.label,
      category: matchedFormat.category,
      extension,
    })

    // Set default target format based on file type
    if (matchedFormat.category === 'image') {
      const targets = IMAGE_CONVERSION_TARGETS[extension] || ['png']
      setTargetFormat(targets[0])
    } else if (extension === 'docx') {
      setTargetFormat('pdf')
    } else if (extension === 'pdf') {
      setTargetFormat('docx')
    } else {
      setTargetFormat('')
    }
  }

  const handleInputChange = (event) => {
    const file = event.target.files[0]
    handleFileProcess(file)
  }

  const handleDragOver = (event) => {
    event.preventDefault()
    setIsDragging(true)
  }

  const handleDragLeave = (event) => {
    event.preventDefault()
    setIsDragging(false)
  }

  const handleDrop = (event) => {
    event.preventDefault()
    setIsDragging(false)
    const file = event.dataTransfer.files[0]
    handleFileProcess(file)
  }

  const handleReset = () => {
    if (convertedResult?.url) {
      URL.revokeObjectURL(convertedResult.url)
    }
    setSelectedFile(null)
    setFileDetails(null)
    setErrorMessage('')
    setConversionError('')
    setConvertedResult(null)
    setTargetFormat('')
    setIsConverting(false)
  }

  // Native Canvas-based image conversion
  const handleConvertImage = () => {
    if (!selectedFile || fileDetails?.category !== 'image') return

    setIsConverting(true)
    setConversionError('')

    const objectUrl = URL.createObjectURL(selectedFile)
    const img = new Image()

    img.onload = () => {
      try {
        const canvas = document.createElement('canvas')
        canvas.width = img.naturalWidth
        canvas.height = img.naturalHeight

        const ctx = canvas.getContext('2d')
        if (!ctx) {
          throw new Error('Could not access 2D canvas context.')
        }

        // Fill background with white when converting transparent PNG/WEBP to JPG
        if (targetFormat === 'jpg' || targetFormat === 'jpeg') {
          ctx.fillStyle = '#FFFFFF'
          ctx.fillRect(0, 0, canvas.width, canvas.height)
        }

        ctx.drawImage(img, 0, 0)

        const mimeType = MIME_TYPES[targetFormat] || 'image/png'
        const quality = targetFormat === 'png' ? undefined : 0.92

        canvas.toBlob(
          (blob) => {
            URL.revokeObjectURL(objectUrl)

            if (!blob) {
              setConversionError('Conversion failed: unable to create image blob.')
              setIsConverting(false)
              return
            }

            const baseName = fileDetails.name.substring(
              0,
              fileDetails.name.lastIndexOf('.')
            )
            const convertedFileName = `${baseName}.${targetFormat}`
            const downloadUrl = URL.createObjectURL(blob)

            setConvertedResult({
              name: convertedFileName,
              size: formatFileSize(blob.size),
              url: downloadUrl,
            })
            setIsConverting(false)
          },
          mimeType,
          quality
        )
      } catch (err) {
        URL.revokeObjectURL(objectUrl)
        setConversionError(err.message || 'An error occurred during image conversion.')
        setIsConverting(false)
      }
    }

    img.onerror = () => {
      URL.revokeObjectURL(objectUrl)
      setConversionError('Could not load image. The file may be corrupt or unreadable.')
      setIsConverting(false)
    }

    img.src = objectUrl
  }

  // DOCX to PDF conversion
  const handleConvertDocx = async () => {
    if (!selectedFile || fileDetails?.extension !== 'docx') return

    setIsConverting(true)
    setConversionError('')
    setConvertedResult(null)

    // Create offscreen container for docx-preview to render DOM
    const tempContainer = document.createElement('div')
    tempContainer.className = 'docx-render-scratchpad'
    tempContainer.style.position = 'absolute'
    tempContainer.style.left = '-9999px'
    tempContainer.style.top = '0'
    tempContainer.style.width = '850px'
    tempContainer.style.background = '#ffffff'
    tempContainer.style.zIndex = '-9999'
    document.body.appendChild(tempContainer)

    try {
      // 1. Read file as ArrayBuffer and render into offscreen DOM
      const arrayBuffer = await selectedFile.arrayBuffer()
      await renderAsync(arrayBuffer, tempContainer, null, {
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
      })

      // 2. Identify rendered sections/pages
      let sections = tempContainer.querySelectorAll(
        '.docx-wrapper > section.docx, section.docx'
      )
      if (!sections || sections.length === 0) {
        const fallback = tempContainer.querySelector('.docx-wrapper') || tempContainer
        sections = [fallback]
      }

      // 3. Initialize jsPDF (A4 standard: 210mm x 297mm)
      const pdf = new jsPDF('p', 'mm', 'a4')
      const pdfWidth = pdf.internal.pageSize.getWidth()
      const pdfHeight = pdf.internal.pageSize.getHeight()

      let isFirstPage = true

      // 4. Capture each rendered section/page
      for (let i = 0; i < sections.length; i++) {
        const section = sections[i]

        const canvas = await html2canvas(section, {
          scale: 2,
          useCORS: true,
          logging: false,
          backgroundColor: '#ffffff',
        })

        const canvasWidth = canvas.width
        const canvasHeight = canvas.height
        const contentHeightInPdf = (canvasHeight * pdfWidth) / canvasWidth

        // If section fits roughly within a single page
        if (contentHeightInPdf <= pdfHeight * 1.05) {
          if (!isFirstPage) {
            pdf.addPage()
          }
          const imgData = canvas.toDataURL('image/jpeg', 0.95)
          pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, contentHeightInPdf)
          isFirstPage = false
        } else {
          // If section is long, slice into multiple A4 pages
          const pageCanvasHeight = Math.floor((canvasWidth * pdfHeight) / pdfWidth)
          let renderedHeight = 0

          while (renderedHeight < canvasHeight) {
            const sliceHeight = Math.min(pageCanvasHeight, canvasHeight - renderedHeight)
            const sliceCanvas = document.createElement('canvas')
            sliceCanvas.width = canvasWidth
            sliceCanvas.height = sliceHeight

            const sCtx = sliceCanvas.getContext('2d')
            sCtx.fillStyle = '#ffffff'
            sCtx.fillRect(0, 0, canvasWidth, sliceHeight)
            sCtx.drawImage(
              canvas,
              0,
              renderedHeight,
              canvasWidth,
              sliceHeight,
              0,
              0,
              canvasWidth,
              sliceHeight
            )

            const sliceImgData = sliceCanvas.toDataURL('image/jpeg', 0.95)
            const sliceHeightInPdf = (sliceHeight * pdfWidth) / canvasWidth

            if (!isFirstPage) {
              pdf.addPage()
            }
            pdf.addImage(sliceImgData, 'JPEG', 0, 0, pdfWidth, sliceHeightInPdf)
            isFirstPage = false

            renderedHeight += sliceHeight
          }
        }
      }

      // 5. Generate downloadable PDF
      const pdfBlob = pdf.output('blob')
      const baseName = fileDetails.name.substring(
        0,
        fileDetails.name.lastIndexOf('.')
      )
      const convertedFileName = `${baseName}.pdf`
      const downloadUrl = URL.createObjectURL(pdfBlob)

      setConvertedResult({
        name: convertedFileName,
        size: formatFileSize(pdfBlob.size),
        url: downloadUrl,
      })
    } catch (err) {
      console.error('DOCX conversion failed:', err)
      setConversionError(
        err.message ||
          'Failed to convert DOCX to PDF. The file may be corrupt or contain unsupported formatting.'
      )
    } finally {
      // 6. Clean up temporary DOM elements
      if (tempContainer && tempContainer.parentNode) {
        tempContainer.parentNode.removeChild(tempContainer)
      }
      setIsConverting(false)
    }
  }

  // PDF to Word (.docx) conversion
  const handleConvertPdf = async () => {
    if (!selectedFile || fileDetails?.extension !== 'pdf') return

    setIsConverting(true)
    setConversionError('')
    setConvertedResult(null)

    try {
      const arrayBuffer = await selectedFile.arrayBuffer()
      const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer })
      const pdfDocument = await loadingTask.promise

      const docChildren = []
      let totalExtractedCharacters = 0

      // Extract text page by page
      for (let pageNum = 1; pageNum <= pdfDocument.numPages; pageNum++) {
        const page = await pdfDocument.getPage(pageNum)
        const textContent = await page.getTextContent()

        // Filter and sort items to reconstruct reading order
        const textItems = textContent.items
          .filter((item) => item.str && item.str.trim().length > 0)
          .map((item) => {
            // item.transform is [scaleX, skewY, skewX, scaleY, tx, ty]
            const x = item.transform[4]
            const y = item.transform[5]
            return {
              text: item.str,
              x,
              y,
              height: item.height || 12,
            }
          })

        if (textItems.length > 0) {
          // Sort items by vertical position (y descending - top to bottom), then horizontal (x ascending)
          textItems.sort((a, b) => {
            const yDiff = b.y - a.y
            // If items share roughly the same baseline (within 4px), sort by X
            if (Math.abs(yDiff) < 4) {
              return a.x - b.x
            }
            return yDiff
          })

          // Group items into lines
          const lines = []
          let currentLine = [textItems[0]]

          for (let i = 1; i < textItems.length; i++) {
            const item = textItems[i]
            const prevItem = currentLine[currentLine.length - 1]

            if (Math.abs(item.y - prevItem.y) < 4) {
              currentLine.push(item)
            } else {
              lines.push(currentLine)
              currentLine = [item]
            }
          }
          if (currentLine.length > 0) {
            lines.push(currentLine)
          }

          // Group lines into readable paragraphs based on line spacing
          let currentParagraphLines = []
          let prevLineY = null
          let avgLineSpacing = 16

          for (let i = 0; i < lines.length; i++) {
            const line = lines[i]
            const lineY = line[0].y
            const lineText = line.map((item) => item.text).join(' ')

            if (prevLineY !== null) {
              const gap = Math.abs(prevLineY - lineY)
              // If gap between lines is noticeably larger than average, treat as new paragraph
              if (gap > avgLineSpacing * 1.7) {
                if (currentParagraphLines.length > 0) {
                  const paraText = currentParagraphLines.join(' ')
                  totalExtractedCharacters += paraText.length
                  docChildren.push(
                    new Paragraph({
                      children: [new TextRun(paraText)],
                      spacing: { after: 120 },
                    })
                  )
                  currentParagraphLines = []
                }
              }
            }

            currentParagraphLines.push(lineText)
            prevLineY = lineY
          }

          // Flush remaining paragraph for current page
          if (currentParagraphLines.length > 0) {
            const paraText = currentParagraphLines.join(' ')
            totalExtractedCharacters += paraText.length
            docChildren.push(
              new Paragraph({
                children: [new TextRun(paraText)],
                spacing: { after: 120 },
              })
            )
          }
        }

        // Add page break between pages (except after the last page)
        if (pageNum < pdfDocument.numPages) {
          docChildren.push(
            new Paragraph({
              children: [new PageBreak()],
            })
          )
        }
      }

      // Check if text was found (handle scanned image PDFs without OCR)
      if (totalExtractedCharacters === 0 || docChildren.length === 0) {
        throw new Error(
          'No selectable text found in this PDF. It appears to be a scanned document or image-only PDF. Converting scanned PDFs requires OCR, which is not supported in this version.'
        )
      }

      // Build DOCX document
      const doc = new Document({
        sections: [
          {
            properties: {},
            children: docChildren,
          },
        ],
      })

      // Generate blob and download URL
      const docBlob = await Packer.toBlob(doc)
      const baseName = fileDetails.name.substring(
        0,
        fileDetails.name.lastIndexOf('.')
      )
      const convertedFileName = `${baseName}.docx`
      const downloadUrl = URL.createObjectURL(docBlob)

      setConvertedResult({
        name: convertedFileName,
        size: formatFileSize(docBlob.size),
        url: downloadUrl,
      })
    } catch (err) {
      console.error('PDF to DOCX conversion failed:', err)
      setConversionError(
        err.message ||
          'Failed to extract text from PDF and build Word document. The file may be password protected or corrupted.'
      )
    } finally {
      setIsConverting(false)
    }
  }

  // Available image targets
  const currentTargets =
    fileDetails?.category === 'image'
      ? IMAGE_CONVERSION_TARGETS[fileDetails.extension] || []
      : []

  return (
    <main className="converter-container">
      <header className="converter-header">
        <h1 className="converter-title">file.type</h1>
        <p className="converter-subtitle">
          In-browser file converter for images and documents
        </p>
      </header>

      {/* Upload Dropzone */}
      {!selectedFile ? (
        <div
          className={`dropzone ${isDragging ? 'dropzone-active' : ''}`}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
        >
          <input
            id="file-upload"
            type="file"
            className="file-input-hidden"
            accept=".png,.jpg,.jpeg,.webp,.pdf,.docx"
            onChange={handleInputChange}
          />
          <label htmlFor="file-upload" className="dropzone-content">
            <div className="upload-icon">📂</div>
            <p className="dropzone-prompt">
              <strong>Click to choose a file</strong> or drag and drop
            </p>
            <p className="supported-formats">
              Supported formats: PNG, JPG, WEBP, PDF, DOCX
            </p>
          </label>
        </div>
      ) : (
        /* File Details and Conversion Card */
        <div className="file-card">
          <div className="file-card-header">
            <span className={`file-badge badge-${fileDetails.category}`}>
              {fileDetails.category.toUpperCase()}
            </span>
            <button
              type="button"
              className="remove-btn"
              onClick={handleReset}
              title="Remove file"
              aria-label="Remove file"
            >
              ✕
            </button>
          </div>

          <div className="file-card-details">
            <div className="detail-item">
              <span className="detail-label">File Name</span>
              <span className="detail-value filename-value" title={fileDetails.name}>
                {fileDetails.name}
              </span>
            </div>

            <div className="detail-item">
              <span className="detail-label">Detected Type</span>
              <span className="detail-value">{fileDetails.typeLabel}</span>
            </div>

            <div className="detail-item">
              <span className="detail-label">File Size</span>
              <span className="detail-value">{fileDetails.size}</span>
            </div>
          </div>

          {/* Image Conversion Controls */}
          {fileDetails.category === 'image' && (
            <div className="conversion-panel">
              <div className="format-selection">
                <label htmlFor="target-format" className="format-label">
                  Convert to:
                </label>
                <select
                  id="target-format"
                  className="format-select"
                  value={targetFormat}
                  onChange={(e) => {
                    setTargetFormat(e.target.value)
                    setConvertedResult(null)
                  }}
                  disabled={isConverting}
                >
                  {currentTargets.map((fmt) => (
                    <option key={fmt} value={fmt}>
                      {fmt.toUpperCase()}
                    </option>
                  ))}
                </select>
              </div>

              <div className="action-row">
                <button
                  type="button"
                  className="btn-primary"
                  onClick={handleConvertImage}
                  disabled={isConverting}
                >
                  {isConverting
                    ? 'Converting Image...'
                    : `Convert to ${targetFormat.toUpperCase()}`}
                </button>
              </div>
            </div>
          )}

          {/* DOCX to PDF Conversion Controls */}
          {fileDetails.extension === 'docx' && (
            <div className="conversion-panel">
              <div className="format-selection">
                <span className="format-label">Convert to:</span>
                <span className="format-fixed-badge">PDF Document (.pdf)</span>
              </div>

              <div className="action-row">
                <button
                  type="button"
                  className="btn-primary"
                  onClick={handleConvertDocx}
                  disabled={isConverting}
                >
                  {isConverting ? 'Converting Document to PDF...' : 'Convert to PDF'}
                </button>
              </div>
            </div>
          )}

          {/* PDF to Word Conversion Controls */}
          {fileDetails.extension === 'pdf' && (
            <div className="conversion-panel">
              <div className="format-selection">
                <span className="format-label">Convert to:</span>
                <span className="format-fixed-badge">Word Document (.docx)</span>
              </div>

              <div className="action-row">
                <button
                  type="button"
                  className="btn-primary"
                  onClick={handleConvertPdf}
                  disabled={isConverting}
                >
                  {isConverting
                    ? 'Extracting Text & Generating Word Document...'
                    : 'Convert to Word (.docx)'}
                </button>
              </div>
            </div>
          )}

          {/* Converted Result Download */}
          {convertedResult && (
            <div className="result-card">
              <div className="result-info">
                <span className="result-icon">✅</span>
                <div>
                  <p className="result-name">{convertedResult.name}</p>
                  <p className="result-size">{convertedResult.size}</p>
                </div>
              </div>
              <a
                href={convertedResult.url}
                download={convertedResult.name}
                className="btn-download"
              >
                Download Converted File
              </a>
            </div>
          )}

          {/* Conversion Error */}
          {conversionError && (
            <div className="error-banner" role="alert">
              <span className="error-icon">⚠️</span>
              <span>{conversionError}</span>
            </div>
          )}

          <div className="file-card-actions">
            <button
              type="button"
              className="btn-secondary"
              onClick={handleReset}
              disabled={isConverting}
            >
              Choose Another File
            </button>
          </div>
        </div>
      )}

      {/* Unsupported File Error Banner */}
      {errorMessage && (
        <div className="error-banner" role="alert">
          <span className="error-icon">⚠️</span>
          <span>{errorMessage}</span>
        </div>
      )}
    </main>
  )
}

export default App
