import React, { useState, useRef, useMemo, useEffect } from 'react';
import * as xlsx from 'xlsx';
import {
  UploadCloud, Download, Search, AlertCircle, Activity,
  ExternalLink, Mail, Phone, Loader2, Bookmark, BookmarkCheck,
  List, DatabaseZap, Trash2, Pencil, Check, X, ArrowRight,
  Building2, Sparkles
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Badge } from '@/components/ui/badge';
import { Progress } from '@/components/ui/progress';
import { Card, CardContent } from '@/components/ui/card';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import {
  Table, TableBody, TableCell, TableHead, TableHeader, TableRow,
} from '@/components/ui/table';
import { useToast } from '@/hooks/use-toast';

// ─── Types ──────────────────────────────────────────────────────────────────

type StudyData = {
  nctId: string;
  briefTitle?: string;
  overallStatus?: string;
  leadSponsor?: string;
  collaborators?: string[];
  therapeuticArea?: string;   // derived from conditions[0]
  indication?: string;         // all conditions joined
  interventions?: string[];
  phase?: string;
  enrollmentCount?: number;
  locationCount?: number;
  startDate?: string;
  origin?: string;             // not in API — user-editable
  contactName?: string;
  contactPhone?: string;
  contactEmail?: string;
  error?: boolean;
};

type SavedStudy = StudyData & {
  savedAt: string;
  origin: string;
};

const STORAGE_KEY = 'ct_saved_studies_v1';

// ─── Helpers ─────────────────────────────────────────────────────────────────

function loadSaved(): SavedStudy[] {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
  } catch {
    return [];
  }
}

function persistSaved(list: SavedStudy[]) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
}

function getStatusColor(status?: string) {
  if (!status) return 'bg-gray-100 text-gray-700 border-gray-200';
  const s = status.toUpperCase();
  if (s === 'RECRUITING') return 'bg-emerald-100 text-emerald-800 border-emerald-200';
  if (s === 'COMPLETED') return 'bg-slate-100 text-slate-700 border-slate-200';
  if (s === 'ACTIVE, NOT RECRUITING') return 'bg-amber-100 text-amber-800 border-amber-200';
  if (s === 'TERMINATED' || s === 'WITHDRAWN' || s === 'SUSPENDED')
    return 'bg-red-100 text-red-800 border-red-200';
  if (s === 'NOT YET RECRUITING') return 'bg-blue-100 text-blue-800 border-blue-200';
  if (s === 'ERROR') return 'bg-destructive/10 text-destructive border-destructive/20';
  return 'bg-gray-100 text-gray-700 border-gray-200';
}

async function fetchStudy(nctId: string): Promise<StudyData> {
  const res = await fetch(
    `https://clinicaltrials.gov/api/v2/studies/${nctId}?fields=NCTId,BriefTitle,OverallStatus,StartDate,LeadSponsorName,CollaboratorName,Condition,InterventionName,InterventionType,Phase,EnrollmentCount,LocationFacility,CentralContactName,CentralContactPhone,CentralContactEMail`
  );
  if (!res.ok) throw new Error('Not ok');
  const json = await res.json();

  const mod = json.protocolSection || {};
  const identification = mod.identificationModule || {};
  const status = mod.statusModule || {};
  const sponsor = mod.sponsorCollaboratorsModule || {};
  const conditions = mod.conditionsModule?.conditions || [];
  const arms = mod.armsInterventionsModule || {};
  const design = mod.designModule || {};
  const contactsMod = mod.contactsLocationsModule || {};
  const centralContacts = contactsMod.centralContacts || [];
  const locations = contactsMod.locations || [];

  // Fallback: if no central contact, pick the first contact from any site
  const siteContact = (() => {
    for (const loc of locations) {
      const c = (loc.contacts || [])[0];
      if (c && (c.name || c.email || c.phone)) return c;
    }
    return null;
  })();
  const contact = centralContacts[0] || siteContact || {};
  const phases: string[] = design.phases || [];
  const interventionList = (arms.interventions || []).map((i: any) => i.name).filter(Boolean);
  const collaboratorList = (sponsor.collaborators || []).map((c: any) => c.name).filter(Boolean);

  return {
    nctId: identification.nctId || nctId,
    briefTitle: identification.briefTitle || '',
    overallStatus: status.overallStatus || '',
    startDate: status.startDateStruct?.date || '',
    leadSponsor: sponsor.leadSponsor?.name || '',
    collaborators: collaboratorList,
    therapeuticArea: conditions[0] || '',
    indication: conditions.join('; '),
    interventions: interventionList,
    phase: phases.join(', ').replace('PHASE', 'Phase'),
    enrollmentCount: design.enrollmentInfo?.count ?? undefined,
    locationCount: locations.length || undefined,
    contactName: contact.name || '',
    contactPhone: contact.phone || '',
    contactEmail: contact.email || '',
    origin: '',
    error: false,
  };
}

// ─── Shared table component ───────────────────────────────────────────────────

function TruncatedCell({ value, maxW = 'max-w-[160px]' }: { value?: string; maxW?: string }) {
  if (!value) return <span className="text-slate-400">-</span>;
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        <span className={`block truncate ${maxW} cursor-default`}>{value}</span>
      </TooltipTrigger>
      <TooltipContent side="bottom" className="max-w-sm leading-relaxed p-3 text-xs">
        {value}
      </TooltipContent>
    </Tooltip>
  );
}

function StudyTable({
  studies,
  isFetching,
  actionSlot,
}: {
  studies: StudyData[];
  isFetching?: boolean;
  actionSlot: (study: StudyData) => React.ReactNode;
}) {
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('All');

  const uniqueStatuses = useMemo(() => {
    const s = new Set(studies.map(s => s.overallStatus).filter((x): x is string => !!x));
    return ['All', ...Array.from(s)];
  }, [studies]);

  const filtered = useMemo(() => studies.filter(s => {
    const q = searchQuery.toLowerCase();
    const matchSearch =
      s.nctId.toLowerCase().includes(q) ||
      (s.leadSponsor || '').toLowerCase().includes(q) ||
      (s.indication || '').toLowerCase().includes(q) ||
      (s.interventions || []).join(' ').toLowerCase().includes(q) ||
      (s.briefTitle || '').toLowerCase().includes(q);
    const matchStatus = statusFilter === 'All' || s.overallStatus === statusFilter;
    return matchSearch && matchStatus;
  }), [studies, searchQuery, statusFilter]);

  return (
    <div className="space-y-3">
      <div className="flex flex-col sm:flex-row gap-3 items-center bg-white dark:bg-slate-900 p-3 rounded-lg border shadow-sm">
        <div className="relative w-full sm:w-80">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-slate-400" />
          <Input
            placeholder="Buscar ensaios, patrocinadores, indicação..."
            value={searchQuery}
            onChange={e => setSearchQuery(e.target.value)}
            className="pl-9 bg-slate-50 dark:bg-slate-950 border-slate-200 dark:border-slate-800"
          />
        </div>
        <div className="flex items-center gap-2 w-full sm:w-auto">
          <span className="text-sm font-medium text-slate-500 whitespace-nowrap">Status:</span>
          <Select value={statusFilter} onValueChange={setStatusFilter}>
            <SelectTrigger className="w-full sm:w-[180px] bg-slate-50 dark:bg-slate-950 border-slate-200 dark:border-slate-800">
              <SelectValue placeholder="Todos" />
            </SelectTrigger>
            <SelectContent>
              {uniqueStatuses.map(s => (
                <SelectItem key={s} value={s}>{s === 'All' ? 'Todos' : s || 'Desconhecido'}</SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>
      </div>

      <div className="bg-white dark:bg-slate-900 rounded-lg border shadow-sm overflow-x-auto">
        <Table>
          <TableHeader className="bg-slate-50/70 dark:bg-slate-950/70">
            <TableRow className="hover:bg-transparent">
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Therapeutic Area</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Sponsor</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Origin (indic.)</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Indication</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Intervention</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Phase</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Status</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap text-right">No. of sites</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap text-right">No. of particip.</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Start</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">NCT ID</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">CRO / Collaborator</TableHead>
              <TableHead className="font-semibold text-slate-800 dark:text-slate-300 whitespace-nowrap">Contact</TableHead>
              <TableHead className="w-[80px]" />
            </TableRow>
          </TableHeader>
          <TableBody>
            {filtered.length === 0 ? (
              <TableRow>
                <TableCell colSpan={14} className="h-40 text-center text-slate-500">
                  {isFetching ? 'Carregando primeiro lote...' : 'Nenhum resultado encontrado.'}
                </TableCell>
              </TableRow>
            ) : (
              filtered.map(study => (
                <TableRow
                  key={study.nctId}
                  className={study.error ? 'bg-red-50/30 dark:bg-red-950/10' : 'hover:bg-slate-50/50 dark:hover:bg-slate-800/30'}
                >
                  {/* Therapeutic Area */}
                  <TableCell className="text-sm"><TruncatedCell value={study.therapeuticArea} maxW="max-w-[140px]" /></TableCell>

                  {/* Sponsor */}
                  <TableCell className="text-sm"><TruncatedCell value={study.leadSponsor} maxW="max-w-[140px]" /></TableCell>

                  {/* Origin (indic.) — not from API */}
                  <TableCell className="text-sm text-slate-400 italic">-</TableCell>

                  {/* Indication */}
                  <TableCell className="text-sm"><TruncatedCell value={study.indication} maxW="max-w-[160px]" /></TableCell>

                  {/* Intervention */}
                  <TableCell className="text-sm">
                    <TruncatedCell value={(study.interventions || []).join('; ')} maxW="max-w-[140px]" />
                  </TableCell>

                  {/* Phase */}
                  <TableCell className="text-sm whitespace-nowrap">
                    {study.phase || <span className="text-slate-400">-</span>}
                  </TableCell>

                  {/* Status */}
                  <TableCell>
                    {study.error ? (
                      <span className="text-destructive text-xs flex items-center gap-1">
                        <AlertCircle className="h-3 w-3" /> Erro
                      </span>
                    ) : (
                      <Badge variant="outline" className={`text-xs font-medium border shadow-none whitespace-nowrap ${getStatusColor(study.overallStatus)}`}>
                        {study.overallStatus || 'Desconhecido'}
                      </Badge>
                    )}
                  </TableCell>

                  {/* No. of sites */}
                  <TableCell className="text-sm text-right tabular-nums">
                    {study.locationCount != null ? study.locationCount : <span className="text-slate-400">-</span>}
                  </TableCell>

                  {/* No. of participants */}
                  <TableCell className="text-sm text-right tabular-nums">
                    {study.enrollmentCount != null ? study.enrollmentCount.toLocaleString() : <span className="text-slate-400">-</span>}
                  </TableCell>

                  {/* Start */}
                  <TableCell className="text-sm whitespace-nowrap tabular-nums">
                    {study.startDate || <span className="text-slate-400">-</span>}
                  </TableCell>

                  {/* NCT ID */}
                  <TableCell className="font-mono text-xs">
                    {study.error ? (
                      <span className="text-slate-500">{study.nctId}</span>
                    ) : (
                      <a
                        href={`https://clinicaltrials.gov/study/${study.nctId}`}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-primary hover:underline underline-offset-4 inline-flex items-center gap-1"
                      >
                        {study.nctId}
                        <ExternalLink className="h-3 w-3 opacity-50" />
                      </a>
                    )}
                  </TableCell>

                  {/* CRO / Collaborator */}
                  <TableCell className="text-sm">
                    <TruncatedCell value={(study.collaborators || []).join('; ')} maxW="max-w-[140px]" />
                  </TableCell>

                  {/* Contact */}
                  <TableCell>
                    {!study.error && (study.contactName || study.contactEmail || study.contactPhone) ? (
                      <div className="space-y-1">
                        {study.contactName && (
                          <div className="text-xs font-medium text-slate-900 dark:text-slate-100">{study.contactName}</div>
                        )}
                        {study.contactEmail && (
                          <a href={`mailto:${study.contactEmail}`} className="text-xs text-primary hover:underline flex items-center gap-1">
                            <Mail className="h-3 w-3 opacity-70" />
                            <span className="truncate max-w-[140px]">{study.contactEmail}</span>
                          </a>
                        )}
                        {study.contactPhone && (
                          <div className="text-xs text-slate-500 flex items-center gap-1">
                            <Phone className="h-3 w-3 opacity-70" />
                            {study.contactPhone}
                          </div>
                        )}
                      </div>
                    ) : (
                      <span className="text-xs text-slate-400">-</span>
                    )}
                  </TableCell>

                  {/* Action */}
                  <TableCell>{actionSlot(study)}</TableCell>
                </TableRow>
              ))
            )}
          </TableBody>
        </Table>
      </div>

      {!isFetching && filtered.length > 0 && (
        <p className="text-center text-xs text-slate-400 py-2">
          {filtered.length} {filtered.length === 1 ? 'estudo' : 'estudos'}
        </p>
      )}
    </div>
  );
}

// ─── Saved list panel ─────────────────────────────────────────────────────────

function SavedList({
  saved,
  onRemove,
  onExport,
  onUpdateOrigin,
}: {
  saved: SavedStudy[];
  onRemove: (nctId: string) => void;
  onExport: () => void;
  onUpdateOrigin: (nctId: string, origin: string) => void;
}) {
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editValue, setEditValue] = useState('');

  // We need to pass edits up — use a ref trick by passing a callback
  // This component is read-only for origin; inline edit handled via prop
  return (
    <div className="space-y-3">
      <div className="flex justify-end gap-2">
        <Button size="sm" variant="outline" onClick={onExport} disabled={saved.length === 0} className="gap-2">
          <Download className="h-4 w-4" /> Exportar CSV
        </Button>
      </div>

      {saved.length === 0 ? (
        <Card className="border-dashed border-2 border-slate-200 dark:border-slate-800 shadow-none">
          <CardContent className="flex flex-col items-center justify-center py-16 text-center">
            <Bookmark className="h-10 w-10 text-slate-300 mb-4" />
            <p className="text-slate-500 font-medium">Nenhum estudo salvo ainda.</p>
            <p className="text-slate-400 text-sm mt-1">Clique em "Salvar" em qualquer estudo da aba Resultados.</p>
          </CardContent>
        </Card>
      ) : (
        <div className="bg-white dark:bg-slate-900 rounded-lg border shadow-sm overflow-x-auto">
          <Table>
            <TableHeader className="bg-slate-50/70 dark:bg-slate-950/70">
              <TableRow className="hover:bg-transparent">
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Therapeutic Area</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Sponsor</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Origin (indic.)</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Indication</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Intervention</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Phase</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Status</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap text-right">No. of sites</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap text-right">No. of particip.</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Start</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">NCT ID</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">CRO / Collaborator</TableHead>
                <TableHead className="font-semibold text-slate-800 whitespace-nowrap">Contact</TableHead>
                <TableHead className="w-[80px]" />
              </TableRow>
            </TableHeader>
            <TableBody>
              {saved.map(study => (
                <TableRow key={study.nctId} className="hover:bg-slate-50/50">
                  <TableCell className="text-sm"><TruncatedCell value={study.therapeuticArea} maxW="max-w-[120px]" /></TableCell>
                  <TableCell className="text-sm"><TruncatedCell value={study.leadSponsor} maxW="max-w-[120px]" /></TableCell>

                  {/* Origin — editable */}
                  <TableCell className="text-sm">
                    {editingId === study.nctId ? (
                      <div className="flex items-center gap-1">
                        <Input
                          value={editValue}
                          onChange={e => setEditValue(e.target.value)}
                          className="h-7 w-28 border-slate-200 bg-white/90 px-2 text-xs"
                          autoFocus
                        />
                        <Button size="icon" variant="ghost" className="h-6 w-6"
                          onClick={() => {
                            onUpdateOrigin(study.nctId, editValue.trim());
                            setEditingId(null);
                          }}>
                          <Check className="h-3 w-3 text-emerald-600" />
                        </Button>
                        <Button size="icon" variant="ghost" className="h-6 w-6"
                          onClick={() => setEditingId(null)}>
                          <X className="h-3 w-3 text-red-500" />
                        </Button>
                      </div>
                    ) : (
                      <span
                        className="flex items-center gap-1 cursor-pointer group"
                        onClick={() => { setEditingId(study.nctId); setEditValue(study.origin || ''); }}
                      >
                        <span className={study.origin ? 'text-slate-700' : 'text-slate-400 italic'}>
                          {study.origin || 'Editar...'}
                        </span>
                        <Pencil className="h-3 w-3 text-slate-300 group-hover:text-slate-500 transition-colors" />
                      </span>
                    )}
                  </TableCell>

                  <TableCell className="text-sm"><TruncatedCell value={study.indication} maxW="max-w-[140px]" /></TableCell>
                  <TableCell className="text-sm"><TruncatedCell value={(study.interventions || []).join('; ')} maxW="max-w-[120px]" /></TableCell>
                  <TableCell className="text-sm whitespace-nowrap">{study.phase || <span className="text-slate-400">-</span>}</TableCell>
                  <TableCell>
                    <Badge variant="outline" className={`text-xs border shadow-none whitespace-nowrap ${getStatusColor(study.overallStatus)}`}>
                      {study.overallStatus || '-'}
                    </Badge>
                  </TableCell>
                  <TableCell className="text-sm text-right tabular-nums">
                    {study.locationCount != null ? study.locationCount : <span className="text-slate-400">-</span>}
                  </TableCell>
                  <TableCell className="text-sm text-right tabular-nums">
                    {study.enrollmentCount != null ? study.enrollmentCount.toLocaleString() : <span className="text-slate-400">-</span>}
                  </TableCell>
                  <TableCell className="text-sm whitespace-nowrap">{study.startDate || <span className="text-slate-400">-</span>}</TableCell>
                  <TableCell className="font-mono text-xs">
                    <a
                      href={`https://clinicaltrials.gov/study/${study.nctId}`}
                      target="_blank" rel="noopener noreferrer"
                      className="text-primary hover:underline underline-offset-4 inline-flex items-center gap-1"
                    >
                      {study.nctId}
                      <ExternalLink className="h-3 w-3 opacity-50" />
                    </a>
                  </TableCell>
                  <TableCell className="text-sm"><TruncatedCell value={(study.collaborators || []).join('; ')} maxW="max-w-[120px]" /></TableCell>
                  <TableCell>
                    {(study.contactName || study.contactEmail || study.contactPhone) ? (
                      <div className="space-y-1">
                        {study.contactName && <div className="text-xs font-medium">{study.contactName}</div>}
                        {study.contactEmail && (
                          <a href={`mailto:${study.contactEmail}`} className="text-xs text-primary hover:underline flex items-center gap-1">
                            <Mail className="h-3 w-3" />
                            <span className="truncate max-w-[120px]">{study.contactEmail}</span>
                          </a>
                        )}
                        {study.contactPhone && (
                          <div className="text-xs text-slate-500 flex items-center gap-1">
                            <Phone className="h-3 w-3" />
                            {study.contactPhone}
                          </div>
                        )}
                      </div>
                    ) : <span className="text-xs text-slate-400">-</span>}
                  </TableCell>
                  <TableCell>
                    <Button
                      size="icon"
                      variant="ghost"
                      className="h-7 w-7 text-slate-400 hover:text-red-500"
                      onClick={() => onRemove(study.nctId)}
                    >
                      <Trash2 className="h-4 w-4" />
                    </Button>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </div>
      )}
    </div>
  );
}

// ─── Main Dashboard ────────────────────────────────────────────────────────────

export default function Dashboard() {
  const { toast } = useToast();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [studies, setStudies] = useState<StudyData[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [isFetching, setIsFetching] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [hasData, setHasData] = useState(false);
  const [activeTab, setActiveTab] = useState('results');
  const [savedStudies, setSavedStudies] = useState<SavedStudy[]>(loadSaved);

  // Persist saved whenever it changes
  useEffect(() => { persistSaved(savedStudies); }, [savedStudies]);

  const savedIds = useMemo(() => new Set(savedStudies.map(s => s.nctId)), [savedStudies]);
  const totalStudies = studies.length + savedStudies.length;
  const sponsorCount = useMemo(() => {
    const sponsors = new Set(
      [...studies, ...savedStudies]
        .map(study => study.leadSponsor?.trim())
        .filter((value): value is string => !!value),
    );
    return sponsors.size;
  }, [studies, savedStudies]);
  const recruitingCount = useMemo(
    () => [...studies, ...savedStudies].filter(study => study.overallStatus?.toUpperCase() === 'RECRUITING').length,
    [studies, savedStudies],
  );
  const highlightedStudies = useMemo(
    () => [...savedStudies, ...studies].filter(study => !!study.contactEmail || !!study.contactPhone).length,
    [savedStudies, studies],
  );

  // ── File upload ──────────────────────────────────────────────────────────────
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsUploading(true);

    try {
      const buffer = await file.arrayBuffer();
      const workbook = xlsx.read(buffer);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1 });

      const nctIds = new Set<string>();
      for (const row of data) {
        if (Array.isArray(row)) {
          for (const cell of row) {
            if (typeof cell === 'string') {
              const match = cell.match(/NCT\d+/i);
              if (match) nctIds.add(match[0].toUpperCase());
            }
          }
        }
      }

      const allIds = Array.from(nctIds);
      if (allIds.length === 0) {
        toast({ title: 'Nenhum ID NCT encontrado no arquivo', variant: 'destructive' });
        return;
      }

      // Filter out IDs already saved
      const newIds = allIds.filter(id => !savedIds.has(id));
      const alreadySaved = allIds.length - newIds.length;

      if (alreadySaved > 0) {
        toast({
          title: `${alreadySaved} estudo(s) já estão na sua lista`,
          description: `Buscando apenas os ${newIds.length} novos.`,
        });
      }

      if (newIds.length === 0) {
        toast({ title: 'Todos os estudos já estão na sua lista salva.' });
        setActiveTab('saved');
        return;
      }

      setHasData(true);
      setActiveTab('results');
      startFetching(newIds);

    } catch (err) {
      console.error(err);
      toast({ title: 'Erro ao processar o arquivo Excel', variant: 'destructive' });
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  // ── Batch fetch ──────────────────────────────────────────────────────────────
  const startFetching = async (nctIds: string[]) => {
    setIsFetching(true);
    setProgress({ current: 0, total: nctIds.length });
    setStudies([]);

    const BATCH_SIZE = 5;
    for (let i = 0; i < nctIds.length; i += BATCH_SIZE) {
      const batch = nctIds.slice(i, i + BATCH_SIZE);
      const results = await Promise.all(
        batch.map(id =>
          fetchStudy(id).catch(() => ({
            nctId: id,
            error: true,
            briefTitle: 'Erro ao carregar',
            overallStatus: 'Erro',
          } as StudyData))
        )
      );
      setStudies(prev => [...prev, ...results]);
      setProgress(prev => ({
        ...prev,
        current: Math.min(prev.current + BATCH_SIZE, prev.total),
      }));
    }
    setIsFetching(false);
  };

  // ── Save / remove ────────────────────────────────────────────────────────────
  const saveStudy = (study: StudyData) => {
    if (savedIds.has(study.nctId)) return;
    const entry: SavedStudy = { ...study, origin: '', savedAt: new Date().toISOString() };
    setSavedStudies(prev => [...prev, entry]);
    toast({ title: `${study.nctId} salvo na sua lista.` });
  };

  const removeStudy = (nctId: string) => {
    setSavedStudies(prev => prev.filter(s => s.nctId !== nctId));
    toast({ title: `${nctId} removido da lista.` });
  };

  const updateSavedOrigin = (nctId: string, origin: string) => {
    setSavedStudies(prev =>
      prev.map(study => (study.nctId === nctId ? { ...study, origin } : study)),
    );
    toast({ title: `${nctId} updated`, description: 'Origin field saved successfully.' });
  };

  // ── CSV export ───────────────────────────────────────────────────────────────
  const escapeField = (v?: string | number | null) => {
    const s = v == null ? '' : String(v);
    return `"${s.replace(/"/g, '""')}"`;
  };

  const exportCSV = (list: StudyData[]) => {
    const headers = [
      'NCT ID', 'Therapeutic Area', 'Sponsor', 'Origin (indic.)', 'Indication',
      'Intervention', 'Phase', 'Status', 'No. of sites', 'No. of participants',
      'Start', 'CRO / Collaborator', 'Contact Name', 'Contact Email', 'Contact Phone',
    ];
    const rows = list.map(d => [
      escapeField(d.nctId),
      escapeField(d.therapeuticArea),
      escapeField(d.leadSponsor),
      escapeField((d as SavedStudy).origin),
      escapeField(d.indication),
      escapeField((d.interventions || []).join('; ')),
      escapeField(d.phase),
      escapeField(d.overallStatus),
      escapeField(d.locationCount),
      escapeField(d.enrollmentCount),
      escapeField(d.startDate),
      escapeField((d.collaborators || []).join('; ')),
      escapeField(d.contactName),
      escapeField(d.contactEmail),
      escapeField(d.contactPhone),
    ]);
    const csv = [headers.map(escapeField).join(','), ...rows.map(r => r.join(','))].join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `clinical_trials_${new Date().toISOString().split('T')[0]}.csv`;
    a.style.visibility = 'hidden';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  // ── Render ───────────────────────────────────────────────────────────────────
  return (
    <div className="min-h-screen bg-[#f6f3ea] px-4 py-4 font-sans md:px-8 md:py-8">
      <div className="mx-auto max-w-[1600px] space-y-6">
        <Input
          type="file"
          accept=".xlsx,.xls"
          className="hidden"
          ref={fileInputRef}
          onChange={handleFileUpload}
        />

        <section className="rounded-[28px] border border-[#e7dfd3] bg-white shadow-[0_18px_44px_rgba(48,26,70,0.06)]">
          <div className="relative overflow-hidden border-b border-[#e7dfd3] px-6 py-6 md:px-8 md:py-8">
            <div className="pointer-events-none absolute inset-0 bg-[radial-gradient(circle_at_top_left,rgba(69,36,107,0.09),transparent_36%),radial-gradient(circle_at_right,rgba(207,128,49,0.14),transparent_28%)]" />
            <div className="relative flex flex-col gap-6 lg:flex-row lg:items-end lg:justify-between">
              <div className="max-w-4xl space-y-5">
                <div className="inline-flex rounded-[26px] border border-[#efe4d8] bg-white/88 px-5 py-4 shadow-[0_14px_30px_rgba(48,26,70,0.05)] backdrop-blur">
                  <img
                    src="/avanzare-logo.jpg"
                    alt="Avanzare Clinical Research Solutions"
                    className="h-auto w-[220px] md:w-[320px]"
                  />
                </div>
                <div className="space-y-3">
                  <div className="text-[11px] font-semibold uppercase tracking-[0.26em] text-[#8d5f37]">
                    Avanzare Clinical Research Solutions
                  </div>
                  <h1 className="max-w-3xl text-3xl font-semibold tracking-[-0.045em] text-[#34204f] md:text-5xl">
                    Clinical trials sourcing workspace
                  </h1>
                  <p className="max-w-2xl text-sm leading-7 text-slate-600 md:text-base">
                    Upload a workbook, detect NCT identifiers, review the live trial data and export a clean working list.
                  </p>
                </div>
              </div>

              <div className="flex items-end lg:justify-end">
                <div className="rounded-2xl border border-[#ecdccc] bg-white/88 px-4 py-3 text-left shadow-[0_10px_24px_rgba(48,26,70,0.04)] backdrop-blur lg:text-right">
                  <div className="text-[11px] font-semibold uppercase tracking-[0.22em] text-[#8d5f37]">
                    By Lisa Palla Tavolaro & Santiago J. Isbert Perlender
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div className="grid gap-0 lg:grid-cols-[1.25fr_0.75fr]">
            <div className="space-y-6 px-6 py-6 md:px-8 md:py-8">
              <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
                {[
                  {
                    label: 'Visible studies',
                    value: totalStudies,
                    note: 'Results and shortlist combined',
                    icon: Activity,
                  },
                  {
                    label: 'Recruiting',
                    value: recruitingCount,
                    note: 'Records ready for prioritization',
                    icon: Sparkles,
                  },
                  {
                    label: 'Sponsors',
                    value: sponsorCount,
                    note: 'Unique organizations in workspace',
                    icon: Building2,
                  },
                  {
                    label: 'With contacts',
                    value: highlightedStudies,
                    note: 'Entries with email or phone',
                    icon: Mail,
                  },
                ].map(metric => (
                  <Card key={metric.label} className="rounded-2xl border-[#eee6dc] bg-white shadow-none">
                    <CardContent className="p-5">
                      <div className="mb-4 flex items-center justify-between">
                        <div className="text-[11px] font-semibold uppercase tracking-[0.22em] text-slate-500">
                          {metric.label}
                        </div>
                        <div className="rounded-full bg-[#45246b]/8 p-2 text-[#45246b]">
                          <metric.icon className="h-4 w-4" />
                        </div>
                      </div>
                      <div className="text-3xl font-semibold tracking-[-0.04em] text-[#34204f]">
                        {metric.value}
                      </div>
                      <p className="mt-2 text-sm leading-6 text-slate-500">{metric.note}</p>
                    </CardContent>
                  </Card>
                ))}
              </div>
            </div>

            <div className="border-t border-[#e7dfd3] bg-[#faf7f2] px-6 py-6 lg:border-l lg:border-t-0 md:px-8 md:py-8">
              <div className="rounded-[24px] border border-[#ece3d8] bg-white p-6 shadow-[0_12px_30px_rgba(48,26,70,0.05)]">
                <div className="mb-6 flex items-start justify-between gap-4">
                  <div>
                    <div className="text-[11px] font-semibold uppercase tracking-[0.22em] text-[#8d5f37]">
                      Upload
                    </div>
                    <h2 className="mt-2 text-2xl font-semibold tracking-[-0.03em] text-[#34204f]">
                      Add a source workbook
                    </h2>
                  </div>
                  <div className="rounded-full bg-[#45246b]/8 p-3 text-[#45246b]">
                    <UploadCloud className="h-5 w-5" />
                  </div>
                </div>

                <div className="space-y-4">
                  <div className="rounded-2xl border border-[#eee6dc] bg-[#fcfbf9] p-4">
                    <div className="text-sm font-medium text-[#34204f]">Expected input</div>
                    <p className="mt-2 text-sm leading-6 text-slate-600">
                      Any `.xlsx` or `.xls` file containing NCT identifiers anywhere in the sheet.
                    </p>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div className="rounded-2xl border border-[#eee6dc] bg-white p-4">
                      <div className="text-[11px] font-semibold uppercase tracking-[0.22em] text-slate-500">Saved</div>
                      <div className="mt-3 text-3xl font-semibold tracking-[-0.04em] text-[#34204f]">{savedStudies.length}</div>
                    </div>
                    <div className="rounded-2xl border border-[#eee6dc] bg-white p-4">
                      <div className="text-[11px] font-semibold uppercase tracking-[0.22em] text-slate-500">Live</div>
                      <div className="mt-3 text-3xl font-semibold tracking-[-0.04em] text-[#34204f]">{studies.length}</div>
                    </div>
                  </div>
                </div>

                <Button
                  size="lg"
                  onClick={() => fileInputRef.current?.click()}
                  disabled={isUploading || isFetching}
                  className="mt-6 h-12 w-full justify-between rounded-2xl bg-[#45246b] px-5 text-sm font-semibold text-white hover:bg-[#341a51]"
                >
                  <span className="flex items-center gap-2">
                    {isUploading ? <Loader2 className="h-4 w-4 animate-spin" /> : <UploadCloud className="h-4 w-4" />}
                    {isUploading ? 'Processing workbook...' : hasData ? 'Upload another workbook' : 'Upload workbook'}
                  </span>
                  <ArrowRight className="h-4 w-4" />
                </Button>
              </div>
            </div>
          </div>
        </section>

        {isFetching && (
          <Card className="overflow-hidden rounded-[26px] border border-[#eadfce] bg-[linear-gradient(135deg,rgba(69,36,107,0.05),rgba(255,255,255,0.88))] shadow-[0_16px_36px_rgba(48,26,70,0.06)]">
            <CardContent className="flex items-center gap-4 p-5">
              <Loader2 className="h-5 w-5 shrink-0 animate-spin text-[#45246b]" />
              <div className="flex-1 space-y-1">
                <div className="flex justify-between text-sm font-medium">
                  <span className="text-[#45246b]">Building your live study set: {progress.current} of {progress.total}</span>
                  <span className="text-[#45246b]/70">{Math.round((progress.current / progress.total) * 100)}%</span>
                </div>
                <Progress value={(progress.current / progress.total) * 100} className="h-2" />
              </div>
            </CardContent>
          </Card>
        )}

        {!hasData && savedStudies.length === 0 && (
          <div className="flex justify-center py-2">
            <Card className="w-full max-w-4xl overflow-hidden rounded-[30px] border border-[#e7dfd3] bg-white shadow-[0_20px_50px_rgba(48,26,70,0.06)]">
              <CardContent className="grid gap-8 p-8 md:grid-cols-[1.1fr_0.9fr] md:p-10">
                <div className="space-y-5">
                  <img
                    src="/avanzare-logo.jpg"
                    alt="Avanzare Clinical Research Solutions"
                    className="h-auto w-[190px]"
                  />
                  <div className="space-y-3">
                    <h3 className="text-3xl font-semibold tracking-[-0.04em] text-[#34204f]">
                      Start with your workbook.
                    </h3>
                    <p className="max-w-xl text-sm leading-7 text-slate-600">
                      Upload an Excel file with NCT identifiers and review the resulting studies in a clean Avanzare workspace.
                    </p>
                  </div>
                  <div className="rounded-[24px] border border-[#ede1ce] bg-[#fbf6ee] p-5">
                    <div className="text-[11px] font-semibold uppercase tracking-[0.22em] text-[#8d5f37]">
                      Workflow
                    </div>
                    <p className="mt-3 text-sm leading-7 text-slate-600">
                      Upload the file, review the fetched trial data, keep the relevant studies and export the final CSV when you are ready.
                    </p>
                  </div>
                </div>
                <div className="flex flex-col justify-center rounded-[26px] border border-[#ece3d8] bg-[#fcfbf8] p-8 text-center">
                  <div className="mx-auto mb-6 flex h-18 w-18 items-center justify-center rounded-full bg-[#45246b]/10">
                    <UploadCloud className="h-9 w-9 text-[#45246b]" />
                  </div>
                  <h3 className="text-xl font-semibold text-[#34204f]">Upload study identifiers</h3>
                  <p className="mx-auto mt-3 max-w-sm text-sm leading-6 text-slate-500">
                    We scan the full workbook, detect NCT IDs automatically, and create a clean study review workspace.
                  </p>
                  <Button size="lg" onClick={() => fileInputRef.current?.click()} disabled={isUploading} className="mt-8 h-12 rounded-2xl bg-[#45246b] px-8 hover:bg-[#341a51]">
                    {isUploading ? <><Loader2 className="mr-2 h-4 w-4 animate-spin" /> Processing...</> : 'Select Excel file'}
                  </Button>
                </div>
              </CardContent>
            </Card>
          </div>
        )}

        {(hasData || savedStudies.length > 0) && (
          <section className="space-y-4">
            <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
              <div>
                <p className="text-xs font-semibold uppercase tracking-[0.2em] text-[#8d5f37]">Workspace</p>
                <h2 className="mt-2 text-2xl font-semibold tracking-[-0.03em] text-[#34204f]">Review, shortlist and export your studies</h2>
                <p className="mt-2 text-sm text-slate-500">
                  Use the results tab for fresh API output and the shortlist tab for the Avanzare follow-up set.
                </p>
              </div>
              <div className="rounded-full border border-[#ece3d8] bg-white/90 px-4 py-2 text-xs font-medium text-slate-500">
                Public registry source • ClinicalTrials.gov • Excel-led workflow
              </div>
            </div>

            <Tabs value={activeTab} onValueChange={setActiveTab}>
              <TabsList className="h-auto rounded-[22px] border border-[#ece3d8] bg-white p-1.5 shadow-[0_10px_30px_rgba(48,26,70,0.05)]">
                <TabsTrigger value="results" className="gap-2 rounded-[16px] px-5 py-3 data-[state=active]:bg-[#45246b] data-[state=active]:text-white data-[state=active]:shadow-none">
                  <DatabaseZap className="h-4 w-4" />
                  Live Results
                  {studies.length > 0 && (
                    <Badge variant="secondary" className="ml-1 bg-white/14 px-1.5 py-0 text-xs text-inherit">{studies.length}</Badge>
                  )}
                </TabsTrigger>
                <TabsTrigger value="saved" className="gap-2 rounded-[16px] px-5 py-3 data-[state=active]:bg-[#45246b] data-[state=active]:text-white data-[state=active]:shadow-none">
                  <List className="h-4 w-4" />
                  Avanzare Shortlist
                  {savedStudies.length > 0 && (
                    <Badge variant="secondary" className="ml-1 bg-white/14 px-1.5 py-0 text-xs text-inherit">{savedStudies.length}</Badge>
                  )}
                </TabsTrigger>
              </TabsList>

              <TabsContent value="results" className="mt-5">
                {studies.length > 0 ? (
                  <>
                    <div className="mb-3 flex justify-end">
                      <Button size="sm" variant="outline" onClick={() => exportCSV(studies)} disabled={studies.length === 0 || isFetching} className="gap-2 rounded-xl border-[#ece3d8] bg-white/90 hover:bg-white">
                        <Download className="h-4 w-4" /> Export CSV
                      </Button>
                    </div>
                    <StudyTable
                      studies={studies}
                      isFetching={isFetching}
                      actionSlot={(study) =>
                        savedIds.has(study.nctId) ? (
                          <Button size="sm" variant="ghost" disabled className="h-7 gap-1 px-2 text-xs text-emerald-600">
                            <BookmarkCheck className="h-3.5 w-3.5" /> Saved
                          </Button>
                        ) : (
                          <Button
                            size="sm"
                            variant="outline"
                            className="h-7 gap-1 rounded-xl border-[#ece3d8] bg-white/90 px-2 text-xs"
                            onClick={() => saveStudy(study)}
                          >
                            <Bookmark className="h-3.5 w-3.5" /> Save
                          </Button>
                        )
                      }
                    />
                  </>
                ) : (
                  !isFetching && (
                    <Card className="mt-4 rounded-[24px] border border-dashed border-[#ddcfbd] bg-white shadow-none">
                      <CardContent className="flex flex-col items-center justify-center py-16 text-slate-500">
                        <UploadCloud className="mb-3 h-8 w-8 text-[#b48654]" />
                        Upload a workbook to populate the live results table.
                      </CardContent>
                    </Card>
                  )
                )}
              </TabsContent>

              <TabsContent value="saved" className="mt-5">
                <SavedList
                  saved={savedStudies}
                  onRemove={removeStudy}
                  onExport={() => exportCSV(savedStudies)}
                  onUpdateOrigin={updateSavedOrigin}
                />
              </TabsContent>
            </Tabs>
          </section>
        )}
      </div>
    </div>
  );
}
