'use client';

import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  collection,
  addDoc,
  deleteDoc,
  doc,
  onSnapshot,
  query,
  where,
  writeBatch,
  getDocs,
  serverTimestamp,
  limit,
} from 'firebase/firestore';
import {
  Upload,
  Barcode,
  Trash2,
  Undo2,
  RotateCcw,
  FileText,
  FileSpreadsheet,
  AlertTriangle,
  CheckCircle2,
  Users,
  Layers,
} from 'lucide-react';
import { db } from '@/lib/firebase';

interface ExpectedProduct {
  ean: string;
  codigoInterno: string;
  referencia: string;
  marca: string;
  preco: number;
  expectedQty: number;
}

interface CompactProduct {
  e: string;
  i: string;
  r: string;
  m: string;
  p: number;
  q: number;
}

interface ScanRecord {
  id: string;
  ean: string;
  scannedBy: string;
  nicheName: string;
  createdAtMs: number;
}

interface NicheHistoryItem {
  niche: string;
  counted: number;
  target: number | null;
}

const SESSION_ID = 'sessao-ativa-loja';
const CHUNK_SIZE = 75;
const BATCH_LIMIT = 3;

function normalizeHeader(h: string) {
  return String(h || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

function normalizeCode(val: string) {
  const cleaned = String(val || '').trim();
  const withoutZeros = cleaned.replace(/^0+/, '');
  return withoutZeros || (cleaned ? '0' : '');
}

function parseNumberBR(val: unknown): number {
  if (typeof val === 'number') return Number.isFinite(val) ? val : 0;
  const cleaned = String(val || '')
    .replace(/\s/g, '')
    .replace('R$', '')
    .replace(/\.(?=\d{3}(?:,|$))/g, '')
    .replace(',', '.');
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function playTone(type: 'ok' | 'error') {
  try {
    const AudioCtx =
      window.AudioContext ||
      (window as unknown as { webkitAudioContext: typeof AudioContext }).webkitAudioContext;
    if (!AudioCtx) return;
    const ctx = new AudioCtx();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);
    if (type === 'ok') {
      osc.frequency.value = 880;
      gain.gain.value = 0.08;
      osc.start();
      osc.stop(ctx.currentTime + 0.08);
    } else {
      osc.type = 'sawtooth';
      osc.frequency.value = 210;
      gain.gain.value = 0.14;
      osc.start();
      osc.stop(ctx.currentTime + 0.28);
    }
  } catch {
    // ignora bloqueio de áudio do navegador
  }
}

export default function InventoryPage() {
  const [step, setStep] = useState<'upload' | 'scan' | 'report'>('upload');
  const [expectedList, setExpectedList] = useState<ExpectedProduct[]>([]);
  const [scans, setScans] = useState<ScanRecord[]>([]);
  const [uploading, setUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState('');
  const [mappedInfo, setMappedInfo] = useState('');

  const [operatorName, setOperatorName] = useState('Operador 1');
  const [currentNiche, setCurrentNiche] = useState('Nicho A');
  const [nicheTargetInput, setNicheTargetInput] = useState('');
  const [nicheTarget, setNicheTarget] = useState<number | null>(null);
  const [nicheCount, setNicheCount] = useState(0);
  const [nicheHistory, setNicheHistory] = useState<NicheHistoryItem[]>([]);

  const [eanInput, setEanInput] = useState('');
  const [alertMsg, setAlertMsg] = useState<{
    type: 'error' | 'warning' | 'success';
    text: string;
  } | null>(null);
  const [brandFilter, setBrandFilter] = useState('TODAS');
  const [statusFilter, setStatusFilter] = useState<
    'all' | 'diff' | 'missing' | 'surplus' | 'ok'
  >('all');

  const scanInputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    const qBlocks = query(
      collection(db, 'inventory_expected_blocks'),
      where('sessionId', '==', SESSION_ID)
    );
    const unsubExpected = onSnapshot(
      qBlocks,
      (snap) => {
        const blocks: { blockIndex: number; items: ExpectedProduct[] }[] = [];
        snap.docs.forEach((d) => {
          const data = d.data();
          if (Array.isArray(data.items)) {
            const expanded: ExpectedProduct[] = (data.items as CompactProduct[]).map((c) => ({
              ean: String(c.e || ''),
              codigoInterno: String(c.i || ''),
              referencia: String(c.r || ''),
              marca: String(c.m || 'Sem Marca'),
              preco: Number(c.p || 0),
              expectedQty: Number(c.q || 0),
            }));
            blocks.push({
              blockIndex: Number(data.blockIndex ?? 0),
              items: expanded,
            });
          }
        });
        blocks.sort((a, b) => a.blockIndex - b.blockIndex);
        setExpectedList(blocks.flatMap((b) => b.items));
      },
      (err) => {
        console.error('Erro no listener de blocos ERP:', err);
        setAlertMsg({
          type: 'error',
          text: `Erro ao ler base ERP no Firestore: ${err.message}`,
        });
      }
    );

    const qScans = query(
      collection(db, 'inventory_scans'),
      where('sessionId', '==', SESSION_ID),
      limit(5000)
    );
    const unsubScans = onSnapshot(
      qScans,
      (snap) => {
        const list: ScanRecord[] = snap.docs
          .map((d) => {
            const data = d.data();
            return {
              id: d.id,
              ean: String(data.ean || ''),
              scannedBy: String(data.scannedBy || ''),
              nicheName: String(data.nicheName || 'Geral'),
              createdAtMs: Number(data.createdAtMs || 0),
            };
          })
          .sort((a, b) => b.createdAtMs - a.createdAtMs);
        setScans(list);
      },
      (err) => {
        console.error('Erro no listener de scans:', err);
        setAlertMsg({
          type: 'error',
          text: `Falha ao ler bipes no Firestore: ${err.message}`,
        });
      }
    );

    return () => {
      unsubExpected();
      unsubScans();
    };
  }, []);

  useEffect(() => {
    if (step === 'scan') {
      scanInputRef.current?.focus();
    }
  }, [step, nicheCount]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    setUploadProgress('Lendo arquivo da planilha...');
    setAlertMsg(null);

    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(buf), { type: 'array', raw: false });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(ws, {
        header: 1,
        raw: false,
        defval: '',
      }) as string[][];

      if (!aoa || aoa.length < 2) {
        setAlertMsg({ type: 'error', text: 'Planilha vazia ou sem dados.' });
        return;
      }

      const headers = aoa[0].map((h, idx) => String(h || '').trim() || `Col_${idx + 1}`);
      const normHeaders = headers.map((h) => normalizeHeader(h));

      const findCol = (exactCandidates: string[], partialCandidates: string[], fallbackIdx: number) => {
        const exactIdx = normHeaders.findIndex((nh) => exactCandidates.includes(nh));
        if (exactIdx !== -1) return exactIdx;
        const partialIdx = normHeaders.findIndex((nh) =>
          partialCandidates.some((pc) => nh.includes(pc))
        );
        if (partialIdx !== -1) return partialIdx;
        return fallbackIdx < headers.length ? fallbackIdx : 0;
      };

      const idxRef = findCol(
        ['referenciamodelo', 'referencia', 'ref', 'modelo'],
        ['referencia', 'modelo'],
        0
      );
      const idxMarca = findCol(
        ['marca', 'brand', 'fabricante', 'grife'],
        ['marca', 'brand', 'fabricante', 'grife'],
        1
      );
      const idxInterno = findCol(
        ['codigointerno', 'codigo', 'codigoproduto', 'sku'],
        ['interno', 'sku'],
        2
      );
      const idxQtd = findCol(
        ['saldodisponivel', 'quantidade', 'qtd', 'estoque', 'saldo'],
        ['saldo', 'qtd', 'estoque'],
        3
      );
      const idxPreco = findCol(
        ['precovenda', 'preco', 'precovendainicial', 'valor'],
        ['precovenda', 'preco', 'valor'],
        5
      );
      const idxEan = findCol(
        ['codigoean', 'ean', 'codigobarras', 'barcode', 'gtin'],
        ['ean', 'barra', 'barcode'],
        6
      );

      setMappedInfo(
        `Colunas detectadas — Ref: "${headers[idxRef]}" | Marca: "${headers[idxMarca]}" | Cód. Interno: "${headers[idxInterno]}" | Saldo: "${headers[idxQtd]}" | Preço: "${headers[idxPreco]}" | EAN: "${headers[idxEan]}"`
      );

      const consolidated = new Map<string, ExpectedProduct>();
      for (let i = 1; i < aoa.length; i++) {
        const row = aoa[i];
        if (!row || row.every((c) => String(c || '').trim() === '')) continue;

        const codigoInterno = String(row[idxInterno] ?? '').replace(/\s+/g, '').trim();
        const ean = String(row[idxEan] ?? '').replace(/\s+/g, '').trim();
        const key = normalizeCode(codigoInterno) || normalizeCode(ean) || `linha_${i + 1}`;

        const referencia = String(row[idxRef] ?? '').trim();
        const marca = String(row[idxMarca] ?? '').trim() || 'Sem Marca';
        const preco = parseNumberBR(row[idxPreco]);
        const qty = Math.trunc(parseNumberBR(row[idxQtd]));

        if (consolidated.has(key)) {
          const prev = consolidated.get(key)!;
          consolidated.set(key, {
            ...prev,
            expectedQty: prev.expectedQty + qty,
            ean: prev.ean || ean,
            codigoInterno: prev.codigoInterno || codigoInterno,
            referencia: prev.referencia || referencia,
            marca: prev.marca !== 'Sem Marca' ? prev.marca : marca,
            preco: prev.preco || preco,
          });
        } else {
          consolidated.set(key, {
            ean,
            codigoInterno,
            referencia,
            marca,
            preco,
            expectedQty: qty,
          });
        }
      }

      const entries = Array.from(consolidated.values());
      const compactEntries: CompactProduct[] = entries.map((item) => ({
        e: item.ean,
        i: item.codigoInterno,
        r: item.referencia,
        m: item.marca,
        p: item.preco,
        q: item.expectedQty,
      }));

      const chunks: CompactProduct[][] = [];
      for (let i = 0; i < compactEntries.length; i += CHUNK_SIZE) {
        chunks.push(compactEntries.slice(i, i + CHUNK_SIZE));
      }

      const oldSnap = await getDocs(
        query(collection(db, 'inventory_expected_blocks'), where('sessionId', '==', SESSION_ID))
      );
      const oldDocs = oldSnap.docs;
      for (let i = 0; i < oldDocs.length; i += 400) {
        const batch = writeBatch(db);
        oldDocs.slice(i, i + 400).forEach((d) => batch.delete(d.ref));
        await batch.commit();
      }

      for (let i = 0; i < chunks.length; i += BATCH_LIMIT) {
        const subChunks = chunks.slice(i, i + BATCH_LIMIT);
        const batch = writeBatch(db);

        subChunks.forEach((items, offset) => {
          const blockIndex = i + offset;
          const docId = `${SESSION_ID}_block_${String(blockIndex).padStart(4, '0')}`;
          const refDoc = doc(db, 'inventory_expected_blocks', docId);
          batch.set(refDoc, {
            sessionId: SESSION_ID,
            blockIndex,
            items,
          });
        });

        await batch.commit();
        const doneCount = Math.min(entries.length, (i + BATCH_LIMIT) * CHUNK_SIZE);
        setUploadProgress(`Sincronizando na nuvem: ${doneCount} de ${entries.length} produtos...`);
      }

      setExpectedList(entries);
      setAlertMsg({
        type: 'success',
        text: `Planilha carregada! ${entries.length} produtos importados em ${chunks.length} blocos.`,
      });
      setStep('scan');
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error('Erro ao importar planilha:', err);
      setAlertMsg({
        type: 'error',
        text: `Erro ao importar planilha: ${msg}`,
      });
    } finally {
      setUploading(false);
      setUploadProgress('');
      e.target.value = '';
    }
  };

  const handleScanSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const code = eanInput.trim();
    if (!code) return;

    if (!/^\d{1,13}$/.test(code)) {
      playTone('error');
      setAlertMsg({
        type: 'error',
        text: `Código inválido ("${code}")! Informe apenas dígitos numéricos (até 13 caracteres). Não foi contabilizado.`,
      });
      setEanInput('');
      scanInputRef.current?.focus();
      return;
    }

    const normalizedCode = normalizeCode(code);
    const existsInErp = expectedList.some(
      (p) =>
        normalizeCode(p.ean) === normalizedCode ||
        normalizeCode(p.codigoInterno) === normalizedCode
    );

    const createdAtMs = Date.now();
    const nicheLabel = (currentNiche || 'Geral').trim();

    try {
      await addDoc(collection(db, 'inventory_scans'), {
        sessionId: SESSION_ID,
        ean: code,
        scannedBy: (operatorName || 'Operador').trim(),
        nicheName: nicheLabel,
        createdAtMs,
        createdAt: serverTimestamp(),
      });

      setNicheCount((c) => c + 1);
      playTone('ok');
      setAlertMsg(
        existsInErp
          ? {
              type: 'success',
              text: `Bipado código ${code} (ref. ${normalizedCode}) em ${nicheLabel}.`,
            }
          : {
              type: 'warning',
              text: `Código válido (${code} / ref. ${normalizedCode}) bipado, mas NÃO consta na planilha do ERP.`,
            }
      );
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error('Erro ao salvar bipe:', err);
      playTone('error');
      setAlertMsg({
        type: 'error',
        text: `Não foi possível salvar no Firestore: ${msg}`,
      });
    }

    setEanInput('');
    scanInputRef.current?.focus();
  };

  const handleDeleteLastScan = async () => {
    const myLast = scans.find((s) => s.scannedBy === operatorName.trim());
    if (!myLast) {
      setAlertMsg({
        type: 'warning',
        text: 'Nenhum bipe recente encontrado para este operador.',
      });
      return;
    }
    await deleteDoc(doc(db, 'inventory_scans', myLast.id));
    if (myLast.nicheName === currentNiche.trim() && nicheCount > 0) {
      setNicheCount((c) => c - 1);
    }
    setAlertMsg({ type: 'success', text: `Último bipe (${myLast.ean}) excluído.` });
    scanInputRef.current?.focus();
  };

  const handleDeleteSingleScan = async (scan: ScanRecord) => {
    await deleteDoc(doc(db, 'inventory_scans', scan.id));
    if (
      scan.scannedBy === operatorName.trim() &&
      scan.nicheName === currentNiche.trim() &&
      nicheCount > 0
    ) {
      setNicheCount((c) => c - 1);
    }
    setAlertMsg({ type: 'success', text: `Bipe ${scan.ean} removido.` });
    scanInputRef.current?.focus();
  };

  const handleFinishAndNextNiche = () => {
    const targetNum = nicheTargetInput.trim()
      ? Math.max(0, parseInt(nicheTargetInput, 10) || 0)
      : nicheTarget;
    if (targetNum !== null && nicheCount !== targetNum) {
      const diff = nicheCount - targetNum;
      const ok = window.confirm(
        `No ${currentNiche} você contou ${targetNum} itens e bipou ${nicheCount} (diferença de ${
          diff > 0 ? `+${diff}` : diff
        }). Deseja fechar este nicho e zerar mesmo assim?`
      );
      if (!ok) return;
    }

    setNicheHistory((prev) => [
      { niche: currentNiche, counted: nicheCount, target: targetNum },
      ...prev,
    ]);

    const trimmed = currentNiche.trim();
    const matchLetter = trimmed.match(/^(.*?)([A-Za-z])$/);
    if (matchLetter) {
      const prefix = matchLetter[1];
      const charCode = matchLetter[2].charCodeAt(0);
      const nextChar = String.fromCharCode(charCode + 1);
      setCurrentNiche(`${prefix}${nextChar}`);
    } else {
      setCurrentNiche('');
    }

    setNicheCount(0);
    setNicheTarget(null);
    setNicheTargetInput('');
    scanInputRef.current?.focus();
  };

  const totalExpectedQty = useMemo(
    () => expectedList.reduce((acc, item) => acc + item.expectedQty, 0),
    [expectedList]
  );

  const totalScannedQty = scans.length;
  const progressPct =
    totalExpectedQty > 0 ? Math.min(100, (totalScannedQty / totalExpectedQty) * 100) : 0;

  const scanCountByNormCode = useMemo(() => {
    const map = new Map<string, number>();
    for (const s of scans) {
      const key = normalizeCode(s.ean);
      if (!key) continue;
      map.set(key, (map.get(key) || 0) + 1);
    }
    return map;
  }, [scans]);

  const consolidatedReport = useMemo(() => {
    const map = new Map<
      string,
      {
        key: string;
        ean: string;
        codigoInterno: string;
        referencia: string;
        marca: string;
        preco: number;
        expectedQty: number;
        scannedQty: number;
      }
    >();

    for (const exp of expectedList) {
      const normInterno = normalizeCode(exp.codigoInterno);
      const normEan = normalizeCode(exp.ean);
      const key = normInterno || normEan || `${exp.referencia}_${exp.marca}`;

      const existing = map.get(key);
      if (existing) {
        existing.expectedQty += exp.expectedQty;
        if (!existing.ean && exp.ean) existing.ean = exp.ean;
        if (!existing.codigoInterno && exp.codigoInterno) existing.codigoInterno = exp.codigoInterno;
      } else {
        const scannedQty =
          (normInterno ? scanCountByNormCode.get(normInterno) || 0 : 0) +
          (normEan && normEan !== normInterno ? scanCountByNormCode.get(normEan) || 0 : 0);

        map.set(key, {
          key,
          ean: exp.ean,
          codigoInterno: exp.codigoInterno,
          referencia: exp.referencia,
          marca: exp.marca,
          preco: exp.preco,
          expectedQty: exp.expectedQty,
          scannedQty,
        });
      }
    }

    for (const [normCode, scannedQty] of scanCountByNormCode.entries()) {
      const alreadyMatched = Array.from(map.values()).some(
        (item) =>
          normalizeCode(item.codigoInterno) === normCode || normalizeCode(item.ean) === normCode
      );
      if (!alreadyMatched) {
        const originalSample = scans.find((s) => normalizeCode(s.ean) === normCode)?.ean || normCode;
        map.set(`unlisted_${normCode}`, {
          key: `unlisted_${normCode}`,
          ean: originalSample,
          codigoInterno: '-',
          referencia: 'NÃO CADASTRADO NO ERP',
          marca: 'Sem Marca',
          preco: 0,
          expectedQty: 0,
          scannedQty,
        });
      }
    }

    return Array.from(map.values()).map((r) => {
      const diff = r.scannedQty - r.expectedQty;
      return {
        ...r,
        diff,
        financialDiff: diff * r.preco,
      };
    });
  }, [expectedList, scans, scanCountByNormCode]);

  const brands = useMemo(() => {
    const set = new Set(consolidatedReport.map((r) => r.marca));
    return ['TODAS', ...Array.from(set).sort()];
  }, [consolidatedReport]);

  const filteredReport = useMemo(() => {
    return consolidatedReport.filter((r) => {
      if (brandFilter !== 'TODAS' && r.marca !== brandFilter) return false;
      if (statusFilter === 'diff') return r.diff !== 0;
      if (statusFilter === 'missing') return r.diff < 0;
      if (statusFilter === 'surplus') return r.diff > 0;
      if (statusFilter === 'ok') return r.diff === 0 && r.expectedQty > 0;
      return true;
    });
  }, [consolidatedReport, brandFilter, statusFilter]);

  const handleDownloadTxt = () => {
    const chronological = [...scans].reverse();
    const content = chronological.map((s) => s.ean).join('\n');
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `codigos-bipados-${new Date().toISOString().slice(0, 10)}.txt`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleDownloadExcel = () => {
    const rows = consolidatedReport.map((r) => ({
      Codigo_Interno: r.codigoInterno,
      EAN: r.ean,
      Referencia: r.referencia,
      Marca: r.marca,
      Preco: r.preco,
      Qtd_ERP: r.expectedQty,
      Qtd_Bipada: r.scannedQty,
      Divergencia: r.diff,
      Impacto_Financeiro: Number(r.financialDiff.toFixed(2)),
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Inventario');
    XLSX.writeFile(wb, `relatorio-inventario-${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 p-4 md:p-8">
      <div className="max-w-7xl mx-auto space-y-6">
        <header className="flex flex-col md:flex-row md:items-center md:justify-between gap-4 border-b border-slate-800 pb-5">
          <div>
            <h1 className="text-2xl font-bold tracking-tight text-white">
              Inventário de Loja — Colaborativo Online
            </h1>
            <p className="text-sm text-slate-400">
              Carga ERP, bipagem EAN-13 ao vivo, contagem por nicho, divergências e exportação TXT.
            </p>
          </div>

          <div className="flex flex-wrap gap-2">
            <button
              onClick={() => setStep('upload')}
              className={`px-4 py-2 rounded-lg text-sm font-medium transition ${
                step === 'upload'
                  ? 'bg-emerald-600 text-white'
                  : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
              }`}
            >
              1. Planilha ERP
            </button>
            <button
              onClick={() => setStep('scan')}
              className={`px-4 py-2 rounded-lg text-sm font-medium transition ${
                step === 'scan'
                  ? 'bg-emerald-600 text-white'
                  : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
              }`}
            >
              2. Bipagem ({totalScannedQty})
            </button>
            <button
              onClick={() => setStep('report')}
              className={`px-4 py-2 rounded-lg text-sm font-medium transition ${
                step === 'report'
                  ? 'bg-emerald-600 text-white'
                  : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
              }`}
            >
              3. Relatório & TXT
            </button>
          </div>
        </header>

        <section className="bg-slate-900 border border-slate-800 rounded-xl p-4 space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2 text-sm">
            <div className="flex items-center gap-4">
              <span>
                Produtos na Lista ERP: <strong className="text-white">{expectedList.length}</strong> SKUs /{' '}
                <strong className="text-white">{totalExpectedQty}</strong> un.
              </span>
              <span>
                Total Bipado (Equipe): <strong className="text-emerald-400">{totalScannedQty}</strong> un.
              </span>
            </div>
            <span className="font-semibold text-emerald-400">{progressPct.toFixed(1)}%</span>
          </div>
          <div className="w-full h-3 bg-slate-800 rounded-full overflow-hidden">
            <div
              className="h-full bg-emerald-500 transition-all duration-300"
              style={{ width: `${progressPct}%` }}
            />
          </div>
          {mappedInfo && <p className="text-xs text-slate-400">{mappedInfo}</p>}
        </section>

        {alertMsg && (
          <div
            className={`p-4 rounded-xl border flex items-center gap-3 text-sm ${
              alertMsg.type === 'error'
                ? 'bg-red-950/70 border-red-700 text-red-200'
                : alertMsg.type === 'warning'
                ? 'bg-amber-950/70 border-amber-700 text-amber-200'
                : 'bg-emerald-950/60 border-emerald-700 text-emerald-200'
            }`}
          >
            {alertMsg.type === 'error' ? (
              <AlertTriangle className="w-5 h-5 shrink-0 text-red-400" />
            ) : (
              <CheckCircle2 className="w-5 h-5 shrink-0 text-emerald-400" />
            )}
            <span>{alertMsg.text}</span>
          </div>
        )}

        {step === 'upload' && (
          <section className="bg-slate-900 border border-slate-800 rounded-xl p-8 text-center space-y-4">
            <div className="w-14 h-14 rounded-full bg-emerald-500/10 text-emerald-400 flex items-center justify-center mx-auto">
              <Upload className="w-7 h-7" />
            </div>
            <h2 className="text-xl font-semibold">Passo 1: Enviar Planilha de Produtos do ERP</h2>
            <p className="text-sm text-slate-400 max-w-xl mx-auto">
              Selecione o arquivo (.xlsx ou .csv) contendo Referência/Modelo, Marca, CodigoInterno, SaldoDisponivel,
              Preço Venda e Código EAN. Os itens são enviados em blocos compactos para o Firestore.
            </p>

            {uploadProgress && (
              <p className="text-sm font-medium text-emerald-400">{uploadProgress}</p>
            )}

            <div className="pt-2">
              <label className="inline-flex items-center gap-2 px-6 py-3 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white font-medium cursor-pointer transition">
                <Upload className="w-4 h-4" />
                <span>{uploading ? 'Importando...' : 'Selecionar Planilha (.xlsx / .csv)'}</span>
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  onChange={handleFileUpload}
                  disabled={uploading}
                  className="hidden"
                />
              </label>
            </div>
          </section>
        )}

        {step === 'scan' && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <div className="lg:col-span-2 space-y-6">
              <section className="bg-slate-900 border border-slate-800 rounded-xl p-6 space-y-4">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div className="flex items-center gap-2 text-sm text-slate-300">
                    <Users className="w-4 h-4 text-emerald-400" />
                    <input
                      type="text"
                      value={operatorName}
                      onChange={(e) => setOperatorName(e.target.value)}
                      placeholder="Seu nome"
                      className="bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-white text-sm"
                    />
                  </div>

                  <div className="flex gap-2">
                    <button
                      type="button"
                      onClick={handleDeleteLastScan}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-rose-950/80 hover:bg-rose-900 border border-rose-700/60 text-rose-200 text-xs font-medium transition"
                    >
                      <Undo2 className="w-3.5 h-3.5" />
                      Excluir Último Código
                    </button>
                  </div>
                </div>

                <form onSubmit={handleScanSubmit} className="space-y-3">
                  <label className="block text-sm font-medium text-slate-300">
                    Bipar ou digitar código EAN-13 / Código Interno
                  </label>
                  <div className="flex gap-2">
                    <div className="relative flex-1">
                      <Barcode className="w-5 h-5 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                      <input
                        ref={scanInputRef}
                        type="text"
                        value={eanInput}
                        onChange={(e) => setEanInput(e.target.value)}
                        placeholder="Ex: 7891234567890 ou 1671216"
                        autoFocus
                        className="w-full pl-11 pr-4 py-3 bg-slate-950 border border-slate-700 focus:border-emerald-500 rounded-xl text-white text-lg tracking-wider outline-none"
                      />
                    </div>
                    <button
                      type="submit"
                      className="px-6 py-3 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white font-semibold transition"
                    >
                      Registrar
                    </button>
                  </div>
                </form>
              </section>

              <section className="bg-slate-900 border border-slate-800 rounded-xl p-6 space-y-4">
                <h3 className="text-base font-semibold text-white">Últimos Códigos Bipados</h3>
                <div className="max-h-96 overflow-y-auto divide-y divide-slate-800">
                  {scans.slice(0, 50).map((s) => (
                    <div key={s.id} className="py-2.5 flex items-center justify-between text-sm">
                      <div className="space-x-3">
                        <span className="font-mono font-semibold text-emerald-400">{s.ean}</span>
                        <span className="text-slate-400">• {s.nicheName}</span>
                        <span className="text-slate-500">({s.scannedBy})</span>
                      </div>
                      <button
                        type="button"
                        onClick={() => handleDeleteSingleScan(s)}
                        title="Excluir este bipe"
                        className="p-1.5 rounded-lg text-slate-400 hover:text-rose-400 hover:bg-rose-950/50 transition"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  ))}
                  {scans.length === 0 && (
                    <p className="py-6 text-center text-slate-500 text-sm">
                      Nenhum código bipado ainda.
                    </p>
                  )}
                </div>
              </section>
            </div>

            <aside className="bg-slate-900 border border-slate-800 rounded-xl p-6 space-y-5 h-fit">
              <div className="flex items-center gap-2">
                <Layers className="w-5 h-5 text-emerald-400" />
                <h2 className="text-lg font-semibold">Contagem por Nicho</h2>
              </div>

              <div className="space-y-3">
                <div>
                  <label className="block text-xs text-slate-400 mb-1">Nicho / Estante Atual</label>
                  <input
                    type="text"
                    value={currentNiche}
                    onChange={(e) => setCurrentNiche(e.target.value)}
                    placeholder="Ex: Nicho A"
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-sm text-white"
                  />
                </div>

                <div>
                  <label className="block text-xs text-slate-400 mb-1">
                    Qtd. Contada no Nicho (opcional para conferir)
                  </label>
                  <input
                    type="number"
                    min={0}
                    value={nicheTargetInput}
                    onChange={(e) => {
                      setNicheTargetInput(e.target.value);
                      const v = parseInt(e.target.value, 10);
                      setNicheTarget(Number.isFinite(v) ? v : null);
                    }}
                    placeholder="Ex: 35"
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-lg text-sm text-white"
                  />
                </div>
              </div>

              <div className="bg-slate-950 border border-slate-800 rounded-xl p-4 text-center space-y-1">
                <p className="text-xs uppercase tracking-wider text-slate-400">Bipados neste nicho</p>
                <p className="text-4xl font-bold text-emerald-400">
                  {nicheCount}
                  {nicheTarget !== null && (
                    <span className="text-lg font-normal text-slate-400"> / {nicheTarget}</span>
                  )}
                </p>
                {nicheTarget !== null && (
                  <p
                    className={`text-xs font-medium ${
                      nicheCount === nicheTarget ? 'text-emerald-400' : 'text-amber-400'
                    }`}
                  >
                    {nicheCount === nicheTarget
                      ? 'Quantidade conferida e batendo!'
                      : `Diferença no nicho: ${nicheCount - nicheTarget}`}
                  </p>
                )}
              </div>

              <button
                type="button"
                onClick={handleFinishAndNextNiche}
                className="w-full inline-flex items-center justify-center gap-2 py-2.5 px-4 rounded-xl bg-slate-800 hover:bg-slate-700 text-white text-sm font-semibold transition"
              >
                <RotateCcw className="w-4 h-4" />
                Zerar e Próximo Nicho
              </button>

              {nicheHistory.length > 0 && (
                <div className="border-t border-slate-800 pt-3 space-y-2">
                  <p className="text-xs font-semibold text-slate-400">Nichos concluídos por você</p>
                  <ul className="space-y-1.5 text-xs">
                    {nicheHistory.map((h, idx) => (
                      <li key={idx} className="flex justify-between text-slate-300">
                        <span>{h.niche}</span>
                        <span>
                          {h.counted} bipados {h.target !== null ? `(contado: ${h.target})` : ''}
                        </span>
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </aside>
          </div>
        )}

        {step === 'report' && (
          <section className="bg-slate-900 border border-slate-800 rounded-xl p-6 space-y-6">
            <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
              <div>
                <h2 className="text-xl font-semibold text-white">
                  Relatório de Inventário e Divergências
                </h2>
                <p className="text-sm text-slate-400">
                  Total de linhas bipadas no TXT:{' '}
                  <strong className="text-emerald-400">{scans.length}</strong>
                </p>
              </div>

              <div className="flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={handleDownloadTxt}
                  className="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-sm font-semibold transition"
                >
                  <FileText className="w-4 h-4" />
                  Baixar TXT (Códigos Bipados)
                </button>
                <button
                  type="button"
                  onClick={handleDownloadExcel}
                  className="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-white text-sm font-semibold transition"
                >
                  <FileSpreadsheet className="w-4 h-4" />
                  Baixar Relatório (.xlsx)
                </button>
              </div>
            </div>

            <div className="flex flex-wrap gap-3">
              <select
                value={brandFilter}
                onChange={(e) => setBrandFilter(e.target.value)}
                className="bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white"
              >
                {brands.map((b) => (
                  <option key={b} value={b}>
                    Marca: {b}
                  </option>
                ))}
              </select>

              <select
                value={statusFilter}
                onChange={(e) =>
                  setStatusFilter(e.target.value as 'all' | 'diff' | 'missing' | 'surplus' | 'ok')
                }
                className="bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white"
              >
                <option value="all">Todos os itens</option>
                <option value="diff">Com divergência</option>
                <option value="missing">Faltando no físico (Bipado &lt; ERP)</option>
                <option value="surplus">Sobrando no físico (Bipado &gt; ERP)</option>
                <option value="ok">Conferidos (Sem diferença)</option>
              </select>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-sm">
                <thead>
                  <tr className="border-b border-slate-800 text-slate-400">
                    <th className="py-2.5 px-3">Cód. Interno</th>
                    <th className="py-2.5 px-3">EAN</th>
                    <th className="py-2.5 px-3">Referência</th>
                    <th className="py-2.5 px-3">Marca</th>
                    <th className="py-2.5 px-3 text-right">Preço</th>
                    <th className="py-2.5 px-3 text-right">Qtd. ERP</th>
                    <th className="py-2.5 px-3 text-right">Qtd. Bipada</th>
                    <th className="py-2.5 px-3 text-right">Divergência</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/70">
                  {filteredReport.slice(0, 200).map((r) => (
                    <tr key={r.key} className="hover:bg-slate-800/30">
                      <td className="py-2 px-3 font-mono text-xs text-slate-300">{r.codigoInterno}</td>
                      <td className="py-2 px-3 font-mono">{r.ean || '-'}</td>
                      <td className="py-2 px-3">{r.referencia}</td>
                      <td className="py-2 px-3">{r.marca}</td>
                      <td className="py-2 px-3 text-right">R$ {r.preco.toFixed(2)}</td>
                      <td className="py-2 px-3 text-right">{r.expectedQty}</td>
                      <td className="py-2 px-3 text-right font-semibold text-emerald-400">
                        {r.scannedQty}
                      </td>
                      <td
                        className={`py-2 px-3 text-right font-bold ${
                          r.diff === 0
                            ? 'text-slate-400'
                            : r.diff < 0
                            ? 'text-rose-400'
                            : 'text-amber-400'
                        }`}
                      >
                        {r.diff > 0 ? `+${r.diff}` : r.diff}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        )}
      </div>
    </div>
  );
}
