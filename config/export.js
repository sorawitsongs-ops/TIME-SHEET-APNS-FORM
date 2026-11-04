(function (global) {
  'use strict';

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function buildExcelHtml({ profile, monthYear, rows }) {
    const name = escapeHtml(profile.name);
    const site = escapeHtml(profile.site);
    const position = escapeHtml(profile.position);
    const department = escapeHtml(profile.department);
    const employeeId = escapeHtml(profile.employeeId);

    let html = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="utf-8">
        <!--[if gte mso 9]>
        <xml>
          <x:ExcelWorkbook>
            <x:ExcelWorksheets>
              <x:ExcelWorksheet>
                <x:Name>Timesheet</x:Name>
                <x:WorksheetOptions>
                  <x:DisplayGridlines/>
                </x:WorksheetOptions>
              </x:ExcelWorksheet>
            </x:ExcelWorksheets>
          </x:ExcelWorkbook>
        </xml>
        <![endif]-->
        <style>
          body, table, th, td, span, div, p {
            font-family: 'Calibri', Arial, sans-serif !important;
          }
          table {
            border-collapse: collapse;
            width: 100%;
            font-family: 'Calibri', Arial, sans-serif !important;
            font-size: 11pt;
          }
          th, td {
            border: 0.5pt solid #444444;
            padding: 6px;
            font-family: 'Calibri', Arial, sans-serif !important;
          }
          th {
            background-color: #E2E8F0;
            font-weight: bold;
            text-align: center;
            font-family: 'Calibri', Arial, sans-serif !important;
          }
          .col-text {
            mso-number-format:"\\@";
            text-align: center;
            font-family: 'Calibri', Arial, sans-serif !important;
          }
          .col-desc {
            mso-number-format:"\\@";
            text-align: left;
            vertical-align: top;
            white-space: pre-wrap;
            font-family: 'Calibri', Arial, sans-serif !important;
          }
          .row-merged {
            background-color: #F8FAFC;
            font-weight: bold;
            text-align: center;
            font-family: 'Calibri', Arial, sans-serif !important;
          }
          .sig-cell {
            border: none !important;
            font-size: 11pt;
            font-family: 'Calibri', Arial, sans-serif !important;
          }
        </style>
      </head>
      <body>
        <table>
          <tr><th colspan="9" style="font-size:14pt; padding:10px;">แบบฟอร์มบันทึกเวลาปฏิบัติงาน (Timesheet)</th></tr>
          <tr>
            <td colspan="4"><b>ชื่อ-สกุล:</b> ${name}</td>
            <td colspan="5"><b>ประจำเดือน/ปี:</b> ${escapeHtml(monthYear)}</td>
          </tr>
          <tr>
            <td colspan="4"><b>หน่วยงาน/สังกัด:</b> ${site}</td>
            <td colspan="5"><b>ตำแหน่ง:</b> ${position}</td>
          </tr>
          <tr>
            <td colspan="4"><b>แผนก:</b> ${department}</td>
            <td colspan="5" class="col-text"><b>รหัสพนักงาน:</b> ${employeeId}</td>
          </tr>
          <tr>
            <th rowspan="2" style="width:40pt;">วันที่</th>
            <th rowspan="2" style="width:55pt;">เข้า<br><span style="font-size:9pt; font-weight:normal;">(hh):(mm)</span></th>
            <th rowspan="2" style="width:55pt;">ออก<br><span style="font-size:9pt; font-weight:normal;">(hh):(mm)</span></th>
            <th rowspan="2" style="width:70pt;">ลายเซ็นต์หัวหน้างาน</th>
            <th colspan="2" style="width:110pt;">O.T</th>
            <th rowspan="2" style="width:70pt;">หัวหน้าอนุมัติ</th>
            <th rowspan="2" style="width:280pt;">รายละเอียดงาน</th>
            <th rowspan="2" style="width:80pt;">หมายเหตุ</th>
          </tr>
          <tr>
            <th style="width:55pt;">เริ่มงาน<br><span style="font-size:9pt; font-weight:normal;">(hh):(mm)</span></th>
            <th style="width:55pt;">ออกเวลา<br><span style="font-size:9pt; font-weight:normal;">(hh):(mm)</span></th>
          </tr>
    `;

    rows.forEach(row => {
      if (['weekend', 'holiday', 'sick-leave', 'annual-leave', 'personal-leave', 'other-leave'].includes(row.type)) {
        html += `
          <tr class="row-merged">
            <td class="col-text">${escapeHtml(row.day)}</td>
            <td colspan="8">${escapeHtml(row.mergedText)}</td>
          </tr>
        `;
      } else {
        const cleanDesc = escapeHtml(row.description || '')
          .replace(/\r?\n/g, '<br style="mso-data-placement:same-cell;" />');

        html += `
          <tr>
            <td class="col-text">${escapeHtml(row.day || '')}</td>
            <td class="col-text">${escapeHtml(row.timeIn || '')}</td>
            <td class="col-text">${escapeHtml(row.timeOut || '')}</td>
            <td>${escapeHtml(row.signHead || '')}</td>
            <td class="col-text">${escapeHtml(row.otStart || '')}</td>
            <td class="col-text">${escapeHtml(row.otEnd || '')}</td>
            <td>${escapeHtml(row.signOtApprove || '')}</td>
            <td class="col-desc">${cleanDesc}</td>
            <td>${escapeHtml(row.remarks || '')}</td>
          </tr>
        `;
      }
    });

    html += `
          <tr><td colspan="9" class="sig-cell" style="height:25pt;"></td></tr>
          <tr>
            <td colspan="4" align="center" class="sig-cell">...................................................</td>
            <td class="sig-cell"></td>
            <td colspan="4" align="center" class="sig-cell">...................................................</td>
          </tr>
          <tr>
            <td colspan="4" align="center" class="sig-cell">( ${name} )</td>
            <td class="sig-cell"></td>
            <td colspan="4" align="center" class="sig-cell">( &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; )</td>
          </tr>
          <tr>
            <td colspan="4" align="center" class="sig-cell"><b>พนักงาน</b></td>
            <td class="sig-cell"></td>
            <td colspan="4" align="center" class="sig-cell"><b>ผู้จัดการ/หัวหน้างาน</b></td>
          </tr>
          <tr>
            <td colspan="4" align="center" class="sig-cell">วันที่ :............................................</td>
            <td class="sig-cell"></td>
            <td colspan="4" align="center" class="sig-cell">วันที่ :............................................</td>
          </tr>
        </table>
      </body>
      </html>
    `;

    return html;
  }

  function exportExcel(data) {
    const html = buildExcelHtml(data);
    const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    const fileName = `Timesheet_${data.profile.name}_${data.monthYear}`.replace(/[\\/:*?"<>|]+/g, '_');

    anchor.href = url;
    anchor.download = `${fileName}.xls`;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);
  }

  global.TimesheetExport = {
    exportExcel,
    printPdf: () => global.print()
  };
})(window);
