import { useState, useEffect } from 'react';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { DragDropContext, Droppable, Draggable } from '@hello-pangea/dnd';
import { 
  Calculator, 
  Save, 
  Trash2, 
  History as HistoryIcon,
  Ruler,
  Maximize,
  Layers,
  FileSpreadsheet,
  Download,
  Plus,
  X,
  FileText,
  GripVertical
} from 'lucide-react';

const DEFAULT_ITEMS = [
  "Pek. Galian Tanah Pondasi Batu Kali",
  "Pek. Urugan Pasir Pondasi Batu Kali",
  "Pek. Batu Kosong Pondasi Batu Kali",
  "Pek. Pasangan Pondasi Batu Kali",
  "Pek. Pengurugan Tanah Kembali",
  "Pek. Pengurugan dan Pemadatan Tanah (per 20 cm) 1m3 Tanah Timbunan",
  "Pek. Pondasi Rollag",
  "Pek. Sloof S1 14x20 cm",
  "Pek. Kolom Praktis 12x12 cm",
  "Pek. Balok Anak 10x15 cm",
  "Pek. Balok Anak 12x20 cm",
  "Pek. RingBalok 10x15 cm",
  "Pek. Balok Latei 11x11 cm",
  "Pek. Plat Meja Dapur 8 cm",
  "Pek. Meja Westafel",
  "Pek. Plat Topi-Topi 10 cm",
  "Pek. Plat Dekker"
];

interface InputRow {
  id: string;
  panjang: string;
  lebar: string;
  tinggi: string;
  panjangVal?: number;
  lebarVal?: number;
  tinggiVal?: number;
  rowVolume?: number;
  unit?: string;
}

interface CalculationRecord {
  id: string;
  item: string;
  
  // legacy fields
  panjang?: string;
  lebar?: string;
  tinggi?: string;
  panjangVal?: number;
  lebarVal?: number;
  tinggiVal?: number;
  
  // new array of inputs
  rows?: InputRow[];

  volume: number;
  unit?: string;
  timestamp: number;
  isDivider?: boolean;
}

const evaluateMath = (expr: string): { value: number; error: boolean; empty: boolean } => {
  if (!expr || !expr.trim()) return { value: 0, error: false, empty: true };
  try {
    const cleanExpr = expr.replace(/,/g, '.').replace(/[^0-9+\-*/(). ]/g, '');
    if (!cleanExpr) return { value: 0, error: false, empty: true };
    const result = new Function(`return ${cleanExpr}`)();
    if (typeof result === 'number' && !isNaN(result)) {
      return { value: result, error: false, empty: false };
    }
    return { value: 0, error: true, empty: false };
  } catch (e) {
    return { value: 0, error: true, empty: false };
  }
};

export default function App() {
  const [customItems, setCustomItems] = useState<string[]>([]);
  const allItems = [...DEFAULT_ITEMS, ...customItems];
  const [selectedItem, setSelectedItem] = useState(DEFAULT_ITEMS[0]);
  const [isEditingItems, setIsEditingItems] = useState(false);
  const [bulkItemsText, setBulkItemsText] = useState('');
  const [inputRows, setInputRows] = useState<InputRow[]>([{ id: crypto.randomUUID(), panjang: '', lebar: '', tinggi: '' }]);
  const [history, setHistory] = useState<CalculationRecord[]>([]);
  const [filterItem, setFilterItem] = useState<string>('All');
  const [pekanKe, setPekanKe] = useState<string>('');
  const [isAddingDivider, setIsAddingDivider] = useState(false);
  const [newDividerText, setNewDividerText] = useState('');

  const handleDragEnd = (result: any) => {
    if (!result.destination) return;
    
    // Only allow reordering if filter is 'All'
    if (filterItem !== 'All') return;

    const items = Array.from(history);
    const [reorderedItem] = items.splice(result.source.index, 1);
    items.splice(result.destination.index, 0, reorderedItem);

    setHistory(items);
  };


  useEffect(() => {
    const saved = localStorage.getItem('calc_history');
    if (saved) {
      try {
        setHistory(JSON.parse(saved));
      } catch (e) {
        console.error('Failed to parse history', e);
      }
    }
    
    const savedCustomItems = localStorage.getItem('calc_custom_items');
    if (savedCustomItems) {
      try {
        const parsed = JSON.parse(savedCustomItems);
        if (Array.isArray(parsed)) {
          setCustomItems(parsed);
        }
      } catch (e) {
        console.error('Failed to parse custom items', e);
      }
    }
  }, []);

  useEffect(() => {
    localStorage.setItem('calc_history', JSON.stringify(history));
  }, [history]);

  useEffect(() => {
    localStorage.setItem('calc_custom_items', JSON.stringify(customItems));
  }, [customItems]);

  const handleSaveBulkItems = () => {
    const newItems = bulkItemsText
      .split('\n')
      .map(line => line.trim())
      .filter(line => line.length > 0 && !DEFAULT_ITEMS.includes(line) && !customItems.includes(line));
    
    if (newItems.length > 0) {
      const updated = [...customItems, ...newItems];
      setCustomItems(updated);
      setSelectedItem(newItems[0]); // Select first newly added item
    }
    setIsEditingItems(false);
    setBulkItemsText('');
  };

  const processedRows = inputRows.map(row => {
    const pEval = evaluateMath(row.panjang);
    const lEval = evaluateMath(row.lebar);
    const tEval = evaluateMath(row.tinggi);

    const hasError = pEval.error || lEval.error || tEval.error;
    const isAllEmpty = pEval.empty && lEval.empty && tEval.empty;
    
    // If an input is empty, default to 1 for multiplication purpose, UNLESS all are empty
    const p = pEval.empty ? 1 : pEval.value;
    const l = lEval.empty ? 1 : lEval.value;
    const t = tEval.empty ? 1 : tEval.value;
    
    let filledCount = 0;
    if (!pEval.empty) filledCount++;
    if (!lEval.empty) filledCount++;
    if (!tEval.empty) filledCount++;

    let unit = 'm³';
    if (filledCount === 1) unit = 'm';
    else if (filledCount === 2) unit = 'm²';
    else if (filledCount === 3) unit = 'm³';
    else unit = '-';

    const volume = isAllEmpty || hasError ? 0 : p * l * t;

    return {
      ...row,
      panjangVal: pEval.empty ? 0 : pEval.value,
      lebarVal: lEval.empty ? 0 : lEval.value,
      tinggiVal: tEval.empty ? 0 : tEval.value,
      hasError,
      isEmpty: isAllEmpty,
      pEval, lEval, tEval,
      volume,
      unit
    };
  });

  const hasAnyError = processedRows.some(r => r.hasError);
  const currentVolume = processedRows.reduce((sum, r) => sum + r.volume, 0);
  
  const firstValidRow = processedRows.find(r => !r.isEmpty && !r.hasError);
  const recordUnit = firstValidRow ? firstValidRow.unit : 'm³';

  const handleSave = () => {
    if (currentVolume === 0 || hasAnyError) return;
    
    const newRecord: CalculationRecord = {
      id: crypto.randomUUID(),
      item: selectedItem,
      rows: processedRows.map(r => ({
        id: r.id,
        panjang: r.panjang.trim() || '-',
        lebar: r.lebar.trim() || '-',
        tinggi: r.tinggi.trim() || '-',
        panjangVal: r.panjangVal,
        lebarVal: r.lebarVal,
        tinggiVal: r.tinggiVal,
        rowVolume: r.volume,
        unit: r.unit
      })),
      volume: currentVolume,
      unit: recordUnit,
      timestamp: Date.now()
    };

    setHistory([newRecord, ...history]);
    setInputRows([{ id: crypto.randomUUID(), panjang: '', lebar: '', tinggi: '' }]);
  };

  const handleDelete = (id: string) => {
    setHistory(history.filter(record => record.id !== id));
  };

  const submitDivider = () => {
    if (!newDividerText.trim()) {
      setIsAddingDivider(false);
      return;
    }
    const newRecord: CalculationRecord = {
      id: crypto.randomUUID(),
      item: newDividerText.trim(),
      volume: 0,
      timestamp: Date.now(),
      isDivider: true
    };
    setHistory([newRecord, ...history]);
    setNewDividerText('');
    setIsAddingDivider(false);
  };

  const clearHistory = () => {
    if (confirm('Apakah Anda yakin ingin menghapus semua riwayat?')) {
      setHistory([]);
    }
  };

  const addRow = () => {
    const lastRow = inputRows[inputRows.length - 1];
    // Copy lebar and tinggi from the previous row if it exists
    setInputRows([...inputRows, { 
      id: crypto.randomUUID(), 
      panjang: '', 
      lebar: lastRow ? lastRow.lebar : '', 
      tinggi: lastRow ? lastRow.tinggi : '' 
    }]);
  };

  const removeRow = (id: string) => {
    if (inputRows.length <= 1) {
      setInputRows([{ id: crypto.randomUUID(), panjang: '', lebar: '', tinggi: '' }]);
      return;
    }
    setInputRows(inputRows.filter(r => r.id !== id));
  };

  const updateRow = (id: string, field: keyof InputRow, value: string) => {
    setInputRows(inputRows.map(r => r.id === id ? { ...r, [field]: value } : r));
  };

  const filteredHistory = filterItem === 'All' 
    ? history 
    : history.filter(h => h.item === filterItem);

  const nonDividerHistory = filteredHistory.filter(r => !r.isDivider);
  const totalVolume = nonDividerHistory.reduce((sum, record) => sum + (record.volume || 0), 0);

  const exportToExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Riwayat');

    // Add Headers
    worksheet.columns = [
      { header: 'Tanggal', key: 'tanggal', width: 22 },
      { header: 'Item Pekerjaan', key: 'item', width: 45 },
      { header: 'Panjang (m)', key: 'panjang', width: 25 },
      { header: 'Lebar (m)', key: 'lebar', width: 25 },
      { header: 'Tinggi (m)', key: 'tinggi', width: 25 },
      { header: 'Hasil', key: 'volume', width: 20 },
    ];

    // Style Headers
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

    let currentRow = 2;

    filteredHistory.forEach(record => {
      if (record.isDivider) {
        worksheet.addRow([record.item]);
        worksheet.mergeCells(`A${currentRow}:F${currentRow}`);
        const dividerRow = worksheet.getRow(currentRow);
        dividerRow.font = { bold: true, color: { argb: 'FF78350F' } }; // amber-900
        dividerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        dividerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE68A' } }; // amber-200
        currentRow++;
        return;
      }

      const dateStr = new Date(record.timestamp).toLocaleString('id-ID', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });
      const startRow = currentRow;

      if (record.rows && record.rows.length > 0) {
        record.rows.forEach(r => {
          const pStr = `${r.panjang}${r.panjang !== '-' && String(r.panjang).match(/[+\-*/()]/) ? `\n= ${(r.panjangVal || 0).toLocaleString('id-ID', { maximumFractionDigits: 4 })}` : ''}`;
          const lStr = `${r.lebar}${r.lebar !== '-' && String(r.lebar).match(/[+\-*/()]/) ? `\n= ${(r.lebarVal || 0).toLocaleString('id-ID', { maximumFractionDigits: 4 })}` : ''}`;
          const tStr = `${r.tinggi}${r.tinggi !== '-' && String(r.tinggi).match(/[+\-*/()]/) ? `\n= ${(r.tinggiVal || 0).toLocaleString('id-ID', { maximumFractionDigits: 4 })}` : ''}`;
          
          worksheet.addRow({
            tanggal: dateStr,
            item: record.item,
            panjang: pStr,
            lebar: lStr,
            tinggi: tStr,
            volume: r.rowVolume ? `${r.rowVolume.toLocaleString('id-ID', { maximumFractionDigits: 4 })} ${r.unit || 'm³'}` : ''
          });
          currentRow++;
        });
        
        // Merge date and item cells if multiple rows
        if (record.rows.length > 1) {
          worksheet.mergeCells(`A${startRow}:A${currentRow - 1}`);
          worksheet.mergeCells(`B${startRow}:B${currentRow - 1}`);
          
          // Add a total row for this record
          worksheet.addRow({
            tanggal: '',
            item: '',
            panjang: '',
            lebar: '',
            tinggi: 'Total Hasil:',
            volume: `${record.volume.toLocaleString('id-ID', { maximumFractionDigits: 4 })} ${record.unit || 'm³'}`
          });
          
          const summaryRow = worksheet.getRow(currentRow);
          summaryRow.font = { bold: true };
          summaryRow.getCell(5).alignment = { horizontal: 'right', vertical: 'middle' };
          
          // Light background for total row
          summaryRow.eachCell({ includeEmpty: true }, (cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F6FF' } };
          });
          worksheet.mergeCells(`A${currentRow}:D${currentRow}`);
          
          currentRow++;
        } else {
          // just one row, update volume to total
          worksheet.getCell(`F${startRow}`).value = `${record.volume.toLocaleString('id-ID', { maximumFractionDigits: 4 })} ${record.unit || 'm³'}`;
          worksheet.getCell(`F${startRow}`).font = { bold: true };
          worksheet.getCell(`F${startRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F6FF' } };
        }

      } else {
        // legacy records without rows
        const pStr = `${record.panjang}${record.panjang !== '-' && String(record.panjang).match(/[+\-*/()]/) ? `\n= ${(record.panjangVal || 0).toLocaleString('id-ID', { maximumFractionDigits: 4 })}` : ''}`;
        const lStr = `${record.lebar}${record.lebar !== '-' && String(record.lebar).match(/[+\-*/()]/) ? `\n= ${(record.lebarVal || 0).toLocaleString('id-ID', { maximumFractionDigits: 4 })}` : ''}`;
        const tStr = `${record.tinggi}${record.tinggi !== '-' && String(record.tinggi).match(/[+\-*/()]/) ? `\n= ${(record.tinggiVal || 0).toLocaleString('id-ID', { maximumFractionDigits: 4 })}` : ''}`;

        worksheet.addRow({
          tanggal: dateStr,
          item: record.item,
          panjang: pStr,
          lebar: lStr,
          tinggi: tStr,
          volume: `${record.volume.toLocaleString('id-ID', { maximumFractionDigits: 4 })} ${record.unit || 'm³'}`
        });
        
        worksheet.getCell(`F${currentRow}`).font = { bold: true };
        worksheet.getCell(`F${currentRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F6FF' } };
        currentRow++;
      }
      
      // format alignments and wrapping for rows of this record
      for (let i = startRow; i < currentRow; i++) {
        const r = worksheet.getRow(i);
        r.alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
        r.getCell(1).alignment = { vertical: 'top', horizontal: 'left' };
        r.getCell(2).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
      }
    });

    // Add empty row for spacing
    worksheet.addRow([]);
    currentRow++;

    // Apply borders
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) { // Apply borders to data rows
        row.eachCell({ includeEmpty: true }, (cell) => {
          cell.border = {
            top: {style:'thin', color: {argb: 'FFD4D4D8'}},
            left: {style:'thin', color: {argb: 'FFD4D4D8'}},
            bottom: {style:'thin', color: {argb: 'FFD4D4D8'}},
            right: {style:'thin', color: {argb: 'FFD4D4D8'}}
          };
        });
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), 'riwayat_kalkulator_konstruksi.xlsx');
  };

  const exportToPDF = () => {
    const doc = new jsPDF();
    
    // Title
    doc.setFontSize(18);
    doc.setFont("helvetica", "bold");
    doc.text("Backup Volume", 14, 20);
    
    // Subtitle
    if (pekanKe) {
      doc.setFontSize(14);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(100);
      doc.text(`Backup Volume Pekan Ke ${pekanKe}`, 14, 28);
    }
    
    const tableData: any[][] = [];
    
    filteredHistory.forEach(record => {
      if (record.isDivider) {
        tableData.push([
          { content: record.item, colSpan: 6, styles: { halign: 'center', fillColor: [253, 230, 138], textColor: [120, 53, 15], fontStyle: 'bold' } }
        ]);
        return;
      }

      const dateStr = new Date(record.timestamp).toLocaleString('id-ID', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });
      
      if (record.rows && record.rows.length > 0) {
        record.rows.forEach((r, idx) => {
          tableData.push([
            idx === 0 ? dateStr : '',
            idx === 0 ? record.item : '',
            r.panjang || '-',
            r.lebar || '-',
            r.tinggi || '-',
            r.rowVolume ? `${r.rowVolume.toLocaleString('id-ID', { maximumFractionDigits: 4 })} ${r.unit || 'm³'}` : ''
          ]);
        });
        
        if (record.rows.length > 1) {
             tableData.push([
              '', '', '', '', 'Total:', `${record.volume.toLocaleString('id-ID', { maximumFractionDigits: 4 })} ${record.unit || 'm³'}`
             ]);
        }
      } else {
        tableData.push([
          dateStr,
          record.item,
          record.panjang || '-',
          record.lebar || '-',
          record.tinggi || '-',
          record.volume.toLocaleString('id-ID', { maximumFractionDigits: 4 }) + ` ${record.unit || 'm³'}`
        ]);
      }
    });

    autoTable(doc, {
      startY: pekanKe ? 35 : 28,
      head: [['Tanggal', 'Item Pekerjaan', 'Panjang (m)', 'Lebar (m)', 'Tinggi (m)', 'Hasil']],
      body: tableData,
      theme: 'striped',
      headStyles: {
        fillColor: [16, 185, 129], // emerald-500
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        halign: 'center',
        valign: 'middle'
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252] // slate-50
      },
      styles: {
        fontSize: 9,
        font: 'helvetica',
        cellPadding: 4,
        textColor: [71, 85, 105], // slate-600
        lineWidth: 0, // No borders
      },
      columnStyles: {
        0: { cellWidth: 32 },
        1: { cellWidth: 'auto' },
        2: { cellWidth: 22, halign: 'right' },
        3: { cellWidth: 22, halign: 'right' },
        4: { cellWidth: 22, halign: 'right' },
        5: { cellWidth: 30, halign: 'right', fontStyle: 'bold', textColor: [15, 118, 110] } // teal-700
      },
      didParseCell: function (data) {
        if (data.cell.text && data.cell.text.length && data.cell.text[0] === 'Total:') {
          if (data.column.index >= 4) {
            data.cell.styles.fillColor = [236, 253, 245]; // emerald-50
            data.cell.styles.textColor = [6, 95, 70]; // emerald-800
            data.cell.styles.fontStyle = 'bold';
          }
        }
      }
    });

    doc.save(`Backup_Volume_${pekanKe ? `Pekan_${pekanKe}` : 'Semua'}.pdf`);
  };

  const renderInputDisplay = (evalResult: { value: number; error: boolean; empty: boolean }) => {
    if (evalResult.empty) return null;
    if (evalResult.error) return <span className="text-rose-500 text-[10px] mt-1 block">Format tidak valid</span>;
    return <span className="text-blue-600 font-medium text-[10px] mt-1 block">= {evalResult.value.toLocaleString('id-ID', { maximumFractionDigits: 4 })}</span>;
  };

  return (
    <div 
      className="min-h-screen text-neutral-900 font-sans selection:bg-blue-200 bg-cover bg-center bg-fixed"
      style={{ backgroundImage: 'url("https://4kwallpapers.com/images/walls/thumbs_3t/26268.jpg")' }}
    >
      <header className="bg-white/80 backdrop-blur-md border-b border-neutral-200/50 sticky top-0 z-10 shadow-sm">
        <div className="w-full mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2 text-blue-700">
            <Calculator className="w-6 h-6" />
            <h1 className="text-xl font-bold tracking-tight">Kalkulator Konstruksi</h1>
          </div>
        </div>
      </header>

      <main className="w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 grid grid-cols-1 lg:grid-cols-12 gap-8 lg:gap-10">
        
        {/* Kalkulator Form */}
        <section className="lg:col-span-4 space-y-6">
          <div className="bg-white/90 backdrop-blur-md rounded-2xl p-6 shadow-xl border border-white/20 flex flex-col h-full">
            <h2 className="text-lg font-semibold mb-6 flex items-center gap-2">
              <FileSpreadsheet className="w-5 h-5 text-blue-600" />
              Input Perhitungan
            </h2>

            <div className="space-y-5 flex-1">
              <div>
                <div className="flex items-center justify-between mb-1.5">
                  <label className="block text-sm font-medium text-neutral-700">
                    Item Pekerjaan
                  </label>
                  {!isEditingItems && (
                    <button 
                      onClick={() => setIsEditingItems(true)}
                      className="text-xs text-blue-600 font-medium hover:text-blue-700 flex items-center gap-1"
                    >
                      <Plus className="w-3 h-3" /> Tambah Item Manual
                    </button>
                  )}
                </div>
                
                {isEditingItems ? (
                  <div className="space-y-2 bg-neutral-50 p-3 rounded-lg border border-neutral-200">
                    <label className="block text-xs font-medium text-neutral-600">
                      Paste daftar item dari Excel (1 item per baris)
                    </label>
                    <textarea
                      value={bulkItemsText}
                      onChange={(e) => setBulkItemsText(e.target.value)}
                      placeholder="Contoh:&#10;Pekerjaan A&#10;Pekerjaan B"
                      rows={5}
                      className="w-full bg-white border border-neutral-300 text-neutral-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block p-2.5 transition-colors resize-none"
                    />
                    <div className="flex gap-2 justify-end">
                      <button
                        onClick={() => {
                          setIsEditingItems(false);
                          setBulkItemsText('');
                        }}
                        className="px-3 py-1.5 text-xs font-medium text-neutral-600 hover:bg-neutral-200 rounded-md transition-colors"
                      >
                        Batal
                      </button>
                      <button
                        onClick={handleSaveBulkItems}
                        className="px-3 py-1.5 text-xs font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-md transition-colors"
                      >
                        Simpan Item
                      </button>
                    </div>
                  </div>
                ) : (
                  <select
                    value={selectedItem}
                    onChange={(e) => setSelectedItem(e.target.value)}
                    className="w-full bg-neutral-50 border border-neutral-300 text-neutral-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block p-2.5 transition-colors"
                  >
                    {allItems.map((item, idx) => (
                      <option key={idx} value={item}>{item}</option>
                    ))}
                  </select>
                )}
              </div>

              <div>
                <div className="flex justify-between items-end mb-2">
                  <label className="block text-sm font-medium text-neutral-700">
                    Dimensi (P x L x T)
                  </label>
                  <span className="text-[10px] text-neutral-400 font-normal">Bisa pakai rumus: 2+3*4</span>
                </div>

                <div className="hidden sm:grid grid-cols-3 gap-3 px-1 mb-1">
                  <label className="text-xs font-medium text-neutral-500 flex items-center gap-1"><Maximize className="w-3 h-3"/> P (m)</label>
                  <label className="text-xs font-medium text-neutral-500 flex items-center gap-1"><Ruler className="w-3 h-3"/> L (m)</label>
                  <label className="text-xs font-medium text-neutral-500 flex items-center gap-1"><Layers className="w-3 h-3"/> T (m)</label>
                </div>

                <div className="space-y-2">
                  {processedRows.map((row) => (
                    <div key={row.id} className="relative group bg-neutral-50 border border-neutral-200 rounded-xl p-3 sm:p-2">
                      {inputRows.length > 1 && (
                        <button 
                          onClick={() => removeRow(row.id)} 
                          className="absolute -top-2 -right-2 bg-white border border-neutral-200 text-neutral-400 hover:text-rose-500 rounded-full p-1 shadow-sm opacity-100 sm:opacity-0 group-hover:opacity-100 transition-opacity"
                          title="Hapus baris"
                        >
                          <X className="w-3 h-3" />
                        </button>
                      )}
                      
                      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                        <div>
                          <label className="sm:hidden block text-xs font-medium text-neutral-500 mb-1">Panjang (m)</label>
                          <input
                            type="text"
                            value={row.panjang}
                            onChange={(e) => updateRow(row.id, 'panjang', e.target.value)}
                            placeholder="Contoh: 1.5 + 2"
                            className="w-full bg-white border border-neutral-300 text-neutral-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block p-2 transition-colors font-mono"
                          />
                          {renderInputDisplay(row.pEval)}
                        </div>
                        <div>
                          <label className="sm:hidden block text-xs font-medium text-neutral-500 mb-1">Lebar (m)</label>
                          <input
                            type="text"
                            value={row.lebar}
                            onChange={(e) => updateRow(row.id, 'lebar', e.target.value)}
                            placeholder="Contoh: 0.14"
                            className="w-full bg-white border border-neutral-300 text-neutral-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block p-2 transition-colors font-mono"
                          />
                          {renderInputDisplay(row.lEval)}
                        </div>
                        <div>
                          <label className="sm:hidden block text-xs font-medium text-neutral-500 mb-1">Tinggi (m)</label>
                          <input
                            type="text"
                            value={row.tinggi}
                            onChange={(e) => updateRow(row.id, 'tinggi', e.target.value)}
                            placeholder="Contoh: 0.20"
                            className="w-full bg-white border border-neutral-300 text-neutral-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block p-2 transition-colors font-mono"
                          />
                          {renderInputDisplay(row.tEval)}
                        </div>
                      </div>
                      {!row.isEmpty && !row.hasError && (
                        <div className="mt-2 text-right text-xs font-semibold text-blue-700 bg-blue-50/50 rounded p-1 border border-blue-100/50">
                          Hasil: {row.volume.toLocaleString('id-ID', { maximumFractionDigits: 4 })} {row.unit}
                        </div>
                      )}
                    </div>
                  ))}
                </div>

                <button 
                  onClick={addRow} 
                  className="mt-3 text-sm text-blue-600 font-medium flex items-center gap-1.5 hover:text-blue-700 transition-colors w-full justify-center py-2 border border-dashed border-blue-200 rounded-lg bg-blue-50/50 hover:bg-blue-50"
                >
                  <Plus className="w-4 h-4" /> Tambah Baris Input
                </button>
              </div>

              <div className="bg-blue-50 rounded-xl p-4 border border-blue-100 flex items-center justify-between mt-auto">
                <div className="w-full">
                  <div className="text-xs font-semibold text-blue-600 uppercase tracking-wider mb-1">Total Saat Ini</div>
                  <div className="text-2xl font-bold text-blue-900 flex justify-between items-baseline">
                    <span className="truncate">
                      {currentVolume > 0 && !hasAnyError ? currentVolume.toLocaleString('id-ID', { maximumFractionDigits: 4 }) : '0'} 
                    </span>
                    <span className="text-sm font-medium text-blue-700 ml-2">{recordUnit}</span>
                  </div>
                </div>
              </div>

              <button
                onClick={handleSave}
                disabled={currentVolume === 0 || hasAnyError}
                className="w-full flex items-center justify-center gap-2 text-white bg-blue-600 hover:bg-blue-700 focus:ring-4 focus:ring-blue-300 font-medium rounded-lg text-sm px-5 py-3 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              >
                <Save className="w-4 h-4" />
                Simpan ke Riwayat
              </button>
            </div>
          </div>
        </section>

        {/* Riwayat Table */}
        <section className="lg:col-span-8 space-y-6">
          <div className="bg-white/90 backdrop-blur-md rounded-2xl shadow-xl border border-white/20 overflow-hidden flex flex-col h-full min-h-[500px]">
            <div className="p-6 object-top border-b border-white/20 bg-white/50 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <h2 className="text-lg font-semibold flex items-center gap-2 text-neutral-800">
                <HistoryIcon className="w-5 h-5 text-neutral-500" />
                Riwayat Perhitungan
              </h2>
              
              <div className="flex flex-wrap items-center gap-2">
                <input
                  type="text"
                  placeholder="Pekan Ke..."
                  value={pekanKe}
                  onChange={(e) => setPekanKe(e.target.value)}
                  className="bg-white border border-neutral-300 text-neutral-700 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 p-2 transition-colors w-24"
                />
                <select
                  value={filterItem}
                  onChange={(e) => setFilterItem(e.target.value)}
                  className="bg-white border border-neutral-300 text-neutral-700 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block p-2 transition-colors min-w-[180px]"
                >
                  <option value="All">Semua Item Pekerjaan</option>
                  {allItems.map((item, idx) => (
                    <option key={idx} value={item}>{item}</option>
                  ))}
                </select>
                
                {filteredHistory.length > 0 && (
                  <button 
                    onClick={exportToExcel}
                    className="flex items-center gap-1.5 px-3 py-2 text-sm font-medium text-emerald-700 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 rounded-lg transition-colors"
                  >
                    <Download className="w-4 h-4" />
                    Excel
                  </button>
                )}

                {filteredHistory.length > 0 && (
                  <button 
                    onClick={exportToPDF}
                    className="flex items-center gap-1.5 px-3 py-2 text-sm font-medium text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-lg transition-colors"
                  >
                    <FileText className="w-4 h-4" />
                    PDF
                  </button>
                )}

                <button 
                  onClick={() => setIsAddingDivider(true)}
                  className="flex items-center gap-1.5 px-3 py-2 text-sm font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded-lg transition-colors"
                >
                  <Layers className="w-4 h-4" />
                  Pembatas
                </button>

                {history.length > 0 && (
                  <button 
                    onClick={clearHistory}
                    className="flex items-center gap-1.5 px-3 py-2 text-sm font-medium text-rose-600 bg-rose-50 hover:bg-rose-100 border border-rose-100 rounded-lg transition-colors"
                  >
                    <Trash2 className="w-4 h-4" />
                    Bersihkan
                  </button>
                )}
              </div>
            </div>

            {isAddingDivider && (
              <div className="px-6 py-3 bg-blue-50/50 border-b border-blue-100 flex items-center gap-2">
                <input
                  type="text"
                  placeholder="Nama Pembatas (Contoh: PEKERJAAN PASANGAN BATA...)"
                  value={newDividerText}
                  onChange={(e) => setNewDividerText(e.target.value)}
                  className="flex-1 bg-white border border-blue-200 text-neutral-800 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 p-2"
                  autoFocus
                  onKeyDown={(e) => e.key === 'Enter' && submitDivider()}
                />
                <button
                  onClick={submitDivider}
                  className="px-4 py-2 bg-blue-600 text-white text-sm font-medium rounded-lg hover:bg-blue-700"
                >
                  Simpan
                </button>
                <button
                  onClick={() => setIsAddingDivider(false)}
                  className="px-4 py-2 bg-white border border-neutral-200 text-neutral-600 text-sm font-medium rounded-lg hover:bg-neutral-50"
                >
                  Batal
                </button>
              </div>
            )}

            <div className="flex-1 p-6 flex flex-col gap-3 overflow-y-auto">
              {filteredHistory.length === 0 ? (
                <div className="flex flex-col items-center justify-center h-full text-neutral-400 py-12">
                  <FileSpreadsheet className="w-12 h-12 mb-3 opacity-20" />
                  <p>Belum ada riwayat perhitungan.</p>
                </div>
              ) : (
                <>
                  <div className="hidden lg:grid grid-cols-12 gap-4 items-center bg-emerald-500/15 backdrop-blur-md border border-emerald-500/20 rounded-xl px-5 py-3 text-xs font-semibold text-emerald-800 uppercase tracking-widest shadow-sm">
                    <div className="col-span-2">Tanggal</div>
                    <div className="col-span-3">Item Pekerjaan</div>
                    <div className="col-span-4 grid grid-cols-3 gap-2 text-right">
                      <div>P (m)</div>
                      <div>L (m)</div>
                      <div>T (m)</div>
                    </div>
                    <div className="col-span-2 text-right text-emerald-900 font-bold">Hasil</div>
                    <div className="col-span-1 text-center">Aksi</div>
                  </div>
                  
                  <DragDropContext onDragEnd={handleDragEnd}>
                    <Droppable droppableId="history-list">
                      {(provided) => (
                        <div {...provided.droppableProps} ref={provided.innerRef} className="space-y-3">
                          {filteredHistory.map((record, index) => (
                            <Draggable key={record.id} draggableId={record.id} index={index} isDragDisabled={filterItem !== 'All'}>
                              {(provided, snapshot) => (
                                <div 
                                  ref={provided.innerRef}
                                  {...provided.draggableProps}
                                  style={{
                                    ...provided.draggableProps.style,
                                    opacity: snapshot.isDragging ? 0.8 : 1
                                  }}
                                  className={`grid grid-cols-1 lg:grid-cols-12 gap-3 lg:gap-4 items-start ${record.isDivider ? 'bg-amber-100/80 border-amber-200' : 'bg-white/70 border-white/50 hover:bg-white/90'} backdrop-blur-md border hover:shadow-md transition-all rounded-xl px-5 py-4 shadow-sm group relative`}
                                >
                                  {/* Drag Handle */}
                                  {filterItem === 'All' && (
                                    <div 
                                      {...provided.dragHandleProps}
                                      className="absolute left-1 top-1/2 -translate-y-1/2 p-1 text-neutral-400 hover:text-neutral-600 cursor-grab active:cursor-grabbing opacity-0 group-hover:opacity-100 transition-opacity"
                                    >
                                      <GripVertical className="w-4 h-4" />
                                    </div>
                                  )}
                                  
                                  {record.isDivider ? (
                                     <>
                                      <div className="col-span-1 lg:col-span-11 flex items-center justify-center font-bold text-lg text-amber-900 tracking-wide text-center">
                                        {record.item}
                                      </div>
                                      <div className="col-span-1 lg:col-span-1 flex justify-end lg:justify-center items-center h-full">
                                        <button
                                          onClick={() => handleDelete(record.id)}
                                          className="text-neutral-400 hover:text-rose-600 hover:bg-rose-50 p-2 rounded-lg transition-colors"
                                          title="Hapus"
                                        >
                                          <Trash2 className="w-4 h-4" />
                                        </button>
                                      </div>
                                     </>
                                  ) : (
                                    <>
                                      <div className="col-span-1 lg:col-span-2 text-xs text-neutral-500 font-medium pt-1 pl-4">
                                        {new Date(record.timestamp).toLocaleDateString('id-ID', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })}
                                      </div>
                                      
                                      <div className="col-span-1 lg:col-span-3 pt-0.5">
                                        <div className="font-medium text-neutral-900 text-sm leading-snug">{record.item}</div>
                                      </div>
                                      
                                      <div className="col-span-1 lg:col-span-4 flex flex-col gap-2">
                                        {(record.rows || [{
                                          id: 'legacy',
                                          panjang: record.panjang, panjangVal: record.panjangVal,
                                          lebar: record.lebar, lebarVal: record.lebarVal,
                                          tinggi: record.tinggi, tinggiVal: record.tinggiVal
                                        }]).map((r, idx) => (
                                          <div key={r.id || idx} className={`grid grid-cols-3 gap-2 text-right ${idx !== (record.rows?.length || 1) - 1 ? 'pb-2 border-b border-neutral-200/50' : ''}`}>
                                            <div className="font-mono text-xs text-neutral-600 break-words">
                                              {r.panjang}
                                              {r.panjang !== '-' && String(r.panjang).match(/[+\-*/()]/) && (
                                                <span className="text-emerald-600 block mt-0.5 font-semibold">= {r.panjangVal?.toLocaleString('id-ID', { maximumFractionDigits: 4 })}</span>
                                              )}
                                            </div>
                                            <div className="font-mono text-xs text-neutral-600 break-words">
                                              {r.lebar}
                                              {r.lebar !== '-' && String(r.lebar).match(/[+\-*/()]/) && (
                                                <span className="text-emerald-600 block mt-0.5 font-semibold">= {r.lebarVal?.toLocaleString('id-ID', { maximumFractionDigits: 4 })}</span>
                                              )}
                                            </div>
                                            <div className="font-mono text-xs text-neutral-600 break-words">
                                              {r.tinggi}
                                              {r.tinggi !== '-' && String(r.tinggi).match(/[+\-*/()]/) && (
                                                <span className="text-emerald-600 block mt-0.5 font-semibold">= {r.tinggiVal?.toLocaleString('id-ID', { maximumFractionDigits: 4 })}</span>
                                              )}
                                            </div>
                                          </div>
                                        ))}
                                      </div>

                                      <div className="col-span-1 lg:col-span-2 flex flex-col justify-center h-full items-end text-sm font-bold text-emerald-800 bg-emerald-50/60 p-2 rounded-lg border border-emerald-100 mt-2 lg:mt-0">
                                        {record.rows && record.rows.length > 1 && (
                                          <div className="text-[10px] text-emerald-700/60 font-medium mb-1 w-full text-right">
                                            {record.rows.map((r, i) => (
                                              <div key={r.id || i} className={i !== record.rows!.length - 1 ? 'mb-1 pb-1 border-b border-emerald-100/50' : ''}>
                                                {r.rowVolume?.toLocaleString('id-ID', { maximumFractionDigits: 4 })} {r.unit || 'm³'}
                                              </div>
                                            ))}
                                          </div>
                                        )}
                                        <div className={`w-full text-right ${record.rows && record.rows.length > 1 ? 'pt-1.5 border-t border-emerald-200/50' : ''}`}>
                                          {(record.volume || 0).toLocaleString('id-ID', { maximumFractionDigits: 4 })} {record.unit || 'm³'}
                                        </div>
                                      </div>

                                      <div className="col-span-1 lg:col-span-1 flex justify-end lg:justify-center items-center h-full mt-2 lg:mt-0">
                                        <button
                                          onClick={() => handleDelete(record.id)}
                                          className="text-neutral-400 hover:text-rose-600 hover:bg-rose-50 p-2 rounded-lg transition-colors"
                                          title="Hapus"
                                        >
                                          <Trash2 className="w-4 h-4" />
                                        </button>
                                      </div>
                                    </>
                                  )}
                                </div>
                              )}
                            </Draggable>
                          ))}
                          {provided.placeholder}
                        </div>
                      )}
                    </Droppable>
                  </DragDropContext>
                </>
              )}
            </div>
          </div>
        </section>

      </main>
    </div>
  );
}

