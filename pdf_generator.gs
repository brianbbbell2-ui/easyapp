/**
 * PDF Generator for Repair Request System V15
 * 
 * IMPORTANT: After saving this file, please click "Run" on any function
 * to authorize Google Docs access (DocumentApp).
 */

const DOC_TEMPLATE_ID = '17cvalvL0MhzpGGoutMv1W20XcHidxhsauxWFJv2RCWE';

function generateRepairPDF_V3(jobNo, formData) {
  try {
    // 1. Validate File Access and Type
    const templateFile = DriveApp.getFileById(DOC_TEMPLATE_ID);
    if (templateFile.getMimeType() !== MimeType.GOOGLE_DOCS) {
      throw new Error("ไฟล์เทมเพลตต้องเป็น 'Google Docs' เท่านั้น (ไอคอนสีฟ้า) ไม่ใช่ไฟล์ Word (.docx)");
    }

    // 2. Create Temporary Copy
    const folder = DriveApp.getFolderById('1vm8eALNA5cghIPcxZ62ll79D9_8oXSuM'); // PDF storage folder
    const copy = templateFile.makeCopy(`RepairForm_${jobNo}`, folder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    // 3. Prepare Data
    const id = formData.id || '';
    const status = formData.status || '';
    const fix1Detail = formData.fix1Detail || '';
    const fix2Detail = formData.fix2Detail || '';

    // 4. Basic Info Replacement
    body.replaceText('{{jobNo}}', jobNo);
    body.replaceText('{{ID}}', id);
    body.replaceText('{{company}}', formData.company || '');
    body.replaceText('{{department}}', formData.department || '');
    body.replaceText('{{costCenter}}', formData.costCenter || '');
    body.replaceText('{{location}}', formData.location || '');
    body.replaceText('{{assetCode}}', formData.assetCode || '');
    body.replaceText('{{assetName}}', formData.assetName || '');
    body.replaceText('{{serialNo}}', formData.serialNo || '');
    body.replaceText('{{problem}}', formData.problem || '');
    body.replaceText('{{warranty}}', formData.warranty || '');
    body.replaceText('{{requester}}', formData.requester || '');
    body.replaceText('{{ext}}', formData.ext || '');
    body.replaceText('{{dateReceived}}', formData.dateReceived || '');

    // 5. Checkbox & Fix Details Replacement (Flexible Logic)
    // Clear placeholders first to avoid showing them if empty
    body.replaceText('{{Fix1Detail}}', '');
    body.replaceText('{{Fix2Detail}}', '');

    // Reset all potential checkboxes in template to empty square
    const checkboxUnchecked = '☐';
    const checkboxChecked = '☑';
    
    // List of status keywords to find and check
    const statusKeywords = [
      { key: 'แก้ไขได้', label: 'แก้ไขได้' },
      { key: 'ส่งซ่อม รายละเอียด', label: 'ส่งซ่อม รายละเอียด', detail: fix1Detail, placeholder: '{{Fix1Detail}}' },
      { key: 'ไม่ส่งซ่อม เนื่องจาก', label: 'ไม่ส่งซ่อม เนื่องจาก', detail: fix2Detail, placeholder: '{{Fix2Detail}}' }
    ];

    statusKeywords.forEach(item => {
      // If this status is selected
      if (status.includes(item.key)) {
        // Find the text in document and prepend checked box
        body.replaceText(item.key, checkboxChecked + ' ' + item.key);
        
        // Handle Details if applicable
        if (item.detail) {
          // Re-replace the empty space we just cleared or directly put detail
          // We find the label again and put detail after it if placeholder was cleared
          body.replaceText(item.key, item.key + ' ' + item.detail);
        }
      } else {
        // For other statuses, just ensure they have unchecked box
        body.replaceText(item.key, checkboxUnchecked + ' ' + item.key);
      }
    });

    // Fallback for FixDetail if they weren't replaced by keyword logic
    if (status.includes('ส่งซ่อม')) {
      body.replaceText('{{Fix1Detail}}', fix1Detail);
    } else if (status.includes('ไม่ส่งซ่อม')) {
      body.replaceText('{{Fix2Detail}}', fix2Detail);
    }

    // 6. Finalize Document
    doc.saveAndClose();

    // 7. Convert to PDF
    const pdfBlob = copy.getAs(MimeType.PDF);
    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setName(`RepairForm_${jobNo}.pdf`);

    // 8. Cleanup Temporary Copy
    copy.setTrashed(true);

    return pdfFile.getUrl();

  } catch (e) {
    console.error('PDF Generation Error:', e);
    throw new Error(`ระบบสร้าง PDF ขัดข้อง: ${e.message}`);
  }
}
