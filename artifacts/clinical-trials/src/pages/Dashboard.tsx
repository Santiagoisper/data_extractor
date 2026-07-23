import React, { useState, useRef, useMemo } from 'react';
import * as xlsx from 'xlsx';
import { UploadCloud, Download, Search, AlertCircle, FileText, CheckCircle2, Activity, ExternalLink, Mail, Phone, CircleDashed, Loader2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Badge } from '@/components/ui/badge';
import { Progress } from '@/components/ui/progress';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from '@/components/ui/table';
import { useToast } from '@/hooks/use-toast';

type StudyData = {
  nctId: string;
  briefTitle?: string;
  overallStatus?: string;
  leadSponsor?: string;
  lastUpdatePostDate?: string;
  contactName?: string;
  contactPhone?: string;
  contactEmail?: string;
  error?: boolean;
};

export default function Dashboard() {
  const { toast } = useToast();
  const fileInputRef = useRef<HTMLInputElement>(null);
  
  const [studies, setStudies] = useState<StudyData[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [isFetching, setIsFetching] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [hasData, setHasData] = useState(false);
  
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('All');

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploading(true);
    setHasData(false);
    
    try {
      const buffer = await file.arrayBuffer();
      const workbook = xlsx.read(buffer);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      
      const data = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1 });
      const nctIds = new Set<string>();
      
      for (const row of data) {
        if (Array.isArray(row)) {
          for (const cell of row) {
            if (typeof cell === 'string') {
              const match = cell.match(/NCT\d+/i);
              if (match) {
                nctIds.add(match[0].toUpperCase());
              }
            }
          }
        }
      }
      
      const uniqueNctIds = Array.from(nctIds);
      if (uniqueNctIds.length === 0) {
        toast({ title: 'No NCT IDs found in the file', variant: 'destructive' });
        setIsUploading(false);
        return;
      }
      
      setHasData(true);
      startFetching(uniqueNctIds);
      
    } catch (err) {
      console.error(err);
      toast({ title: 'Error parsing Excel file', variant: 'destructive' });
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const startFetching = async (nctIds: string[]) => {
    setIsFetching(true);
    setProgress({ current: 0, total: nctIds.length });
    setStudies([]);
    
    const results: StudyData[] = [];
    const BATCH_SIZE = 5;
    
    for (let i = 0; i < nctIds.length; i += BATCH_SIZE) {
      const batch = nctIds.slice(i, i + BATCH_SIZE);
      
      const promises = batch.map(async (nctId) => {
        try {
          const res = await fetch(`https://clinicaltrials.gov/api/v2/studies/${nctId}?fields=NCTId,BriefTitle,OverallStatus,LeadSponsorName,LastUpdatePostDate,CentralContactName,CentralContactPhone,CentralContactEMail`);
          if (!res.ok) throw new Error('Not ok');
          const json = await res.json();
          
          const mod = json.protocolSection || {};
          const contactsLoc = mod.contactsLocationsModule || {};
          const centralContacts = contactsLoc.centralContacts || [];
          const contact = centralContacts[0] || {};
          
          return {
            nctId: mod.identificationModule?.nctId || nctId,
            briefTitle: mod.identificationModule?.briefTitle || '',
            overallStatus: mod.statusModule?.overallStatus || '',
            leadSponsor: mod.sponsorCollaboratorsModule?.leadSponsor?.name || '',
            lastUpdatePostDate: mod.statusModule?.lastUpdatePostDateStruct?.date || '',
            contactName: contact.name || '',
            contactPhone: contact.phone || '',
            contactEmail: contact.email || '',
            error: false
          };
        } catch (err) {
          return {
            nctId,
            error: true,
            briefTitle: 'Error loading',
            overallStatus: 'Error',
            leadSponsor: 'Error loading',
            lastUpdatePostDate: '',
            contactName: '',
            contactPhone: '',
            contactEmail: '',
          };
        }
      });
      
      const batchResults = await Promise.all(promises);
      results.push(...batchResults);
      
      setStudies(prev => [...prev, ...batchResults]);
      setProgress(prev => ({ ...prev, current: Math.min(prev.current + BATCH_SIZE, prev.total) }));
    }
    
    setIsFetching(false);
  };

  const filteredStudies = useMemo(() => {
    return studies.filter(s => {
      const matchesSearch = 
        s.nctId.toLowerCase().includes(searchQuery.toLowerCase()) ||
        (s.briefTitle || '').toLowerCase().includes(searchQuery.toLowerCase()) ||
        (s.leadSponsor || '').toLowerCase().includes(searchQuery.toLowerCase()) ||
        (s.contactName || '').toLowerCase().includes(searchQuery.toLowerCase());
        
      const matchesStatus = statusFilter === 'All' || s.overallStatus === statusFilter;
      
      return matchesSearch && matchesStatus;
    });
  }, [studies, searchQuery, statusFilter]);

  const uniqueStatuses = useMemo(() => {
    const statuses = new Set(studies.map(s => s.overallStatus).filter((s): s is string => !!s));
    return ['All', ...Array.from(statuses)];
  }, [studies]);

  const exportCSV = () => {
    // RFC 4180-compliant CSV: always wrap every field in double-quotes and
    // escape any internal double-quote by doubling it.  This handles commas,
    // newlines, and quotes inside values such as "ACTIVE, NOT RECRUITING".
    const escapeField = (value: string | null | undefined): string => {
      const str = value == null ? '' : String(value);
      return `"${str.replace(/"/g, '""')}"`;
    };

    const headers = [
      'NCT ID',
      'Study Title',
      'Sponsor',
      'Status',
      'Last Update',
      'Contact Name',
      'Contact Phone',
      'Contact Email',
    ];

    const rows = filteredStudies.map(d => [
      escapeField(d.nctId),
      escapeField(d.briefTitle),
      escapeField(d.leadSponsor),
      escapeField(d.overallStatus),
      escapeField(d.lastUpdatePostDate),
      escapeField(d.contactName),
      escapeField(d.contactPhone),
      escapeField(d.contactEmail),
    ]);

    const csvContent = [
      headers.map(escapeField).join(','),
      ...rows.map(r => r.join(',')),
    ].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', `clinical_trials_export_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const getStatusColor = (status?: string) => {
    if (!status) return 'bg-gray-100 text-gray-800 border-gray-200';
    const s = status.toUpperCase();
    if (s === 'RECRUITING') return 'bg-emerald-100 text-emerald-800 border-emerald-200';
    if (s === 'COMPLETED') return 'bg-slate-100 text-slate-800 border-slate-200';
    if (s === 'ACTIVE, NOT RECRUITING') return 'bg-amber-100 text-amber-800 border-amber-200';
    if (s === 'TERMINATED' || s === 'WITHDRAWN' || s === 'SUSPENDED') return 'bg-red-100 text-red-800 border-red-200';
    if (s === 'NOT YET RECRUITING') return 'bg-blue-100 text-blue-800 border-blue-200';
    if (s === 'ERROR') return 'bg-destructive/10 text-destructive border-destructive/20';
    return 'bg-gray-100 text-gray-800 border-gray-200';
  };

  return (
    <div className="min-h-screen bg-slate-50 dark:bg-slate-950 p-6 md:p-12 font-sans">
      <div className="max-w-[1400px] mx-auto space-y-6">
        
        <header className="flex flex-col md:flex-row md:items-end justify-between gap-4">
          <div>
            <h1 className="text-2xl font-bold tracking-tight text-slate-900 dark:text-slate-50 flex items-center gap-2">
              <Activity className="h-6 w-6 text-primary" />
              Trial Contacts Workbench
            </h1>
            <p className="text-sm text-slate-500 dark:text-slate-400 mt-1">
              Bulk extract and resolve clinical trial sponsor contacts from NCT IDs.
            </p>
          </div>
          
          {hasData && (
            <div className="flex items-center gap-3">
              <Button variant="outline" onClick={() => {
                setStudies([]);
                setHasData(false);
                setSearchQuery('');
                setStatusFilter('All');
              }}>
                Upload New File
              </Button>
              <Button onClick={exportCSV} disabled={filteredStudies.length === 0 || isFetching} className="gap-2 shadow-sm">
                <Download className="h-4 w-4" />
                Export CSV
              </Button>
            </div>
          )}
        </header>

        {!hasData ? (
          <div className="mt-12 flex justify-center">
            <Card className="w-full max-w-2xl border-dashed border-2 border-slate-300 dark:border-slate-800 bg-white/50 dark:bg-slate-900/50 hover:bg-white dark:hover:bg-slate-900 transition-colors shadow-sm">
              <CardContent className="flex flex-col items-center justify-center p-12 text-center h-[400px]">
                <div className="h-16 w-16 rounded-full bg-primary/10 flex items-center justify-center mb-6">
                  <UploadCloud className="h-8 w-8 text-primary" />
                </div>
                <h3 className="text-xl font-semibold text-slate-900 dark:text-slate-50 mb-2">Upload study list</h3>
                <p className="text-slate-500 dark:text-slate-400 mb-8 max-w-sm">
                  Drop an Excel (.xlsx) file containing NCT IDs. The parser will automatically extract any cell matching the NCT format.
                </p>
                <Input 
                  type="file" 
                  accept=".xlsx,.xls" 
                  className="hidden" 
                  ref={fileInputRef}
                  onChange={handleFileUpload}
                />
                <Button 
                  size="lg" 
                  onClick={() => fileInputRef.current?.click()}
                  disabled={isUploading}
                  className="shadow-sm font-medium px-8"
                >
                  {isUploading ? (
                    <><Loader2 className="mr-2 h-4 w-4 animate-spin" /> Parsing...</>
                  ) : (
                    'Select Excel File'
                  )}
                </Button>
              </CardContent>
            </Card>
          </div>
        ) : (
          <div className="space-y-4 animate-in fade-in slide-in-from-bottom-4 duration-500">
            {isFetching && (
              <Card className="border-primary/20 bg-primary/5 shadow-sm">
                <CardContent className="p-4 flex items-center gap-4">
                  <Loader2 className="h-5 w-5 text-primary animate-spin" />
                  <div className="flex-1 space-y-1">
                    <div className="flex justify-between text-sm font-medium">
                      <span className="text-primary">Fetching {progress.current} of {progress.total} studies...</span>
                      <span className="text-primary/70">{Math.round((progress.current / progress.total) * 100)}%</span>
                    </div>
                    <Progress value={(progress.current / progress.total) * 100} className="h-2" />
                  </div>
                </CardContent>
              </Card>
            )}

            <div className="flex flex-col sm:flex-row gap-4 justify-between items-center bg-white dark:bg-slate-900 p-4 rounded-lg border shadow-sm">
              <div className="relative w-full sm:w-80">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-slate-400" />
                <Input 
                  placeholder="Search trials, sponsors, contacts..." 
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="pl-9 bg-slate-50 dark:bg-slate-950 border-slate-200 dark:border-slate-800 focus-visible:ring-primary/20"
                />
              </div>
              <div className="flex items-center gap-2 w-full sm:w-auto">
                <span className="text-sm font-medium text-slate-500 whitespace-nowrap">Filter Status:</span>
                <Select value={statusFilter} onValueChange={setStatusFilter}>
                  <SelectTrigger className="w-full sm:w-[180px] bg-slate-50 dark:bg-slate-950 border-slate-200 dark:border-slate-800">
                    <SelectValue placeholder="All Statuses" />
                  </SelectTrigger>
                  <SelectContent>
                    {uniqueStatuses.map(status => (
                      <SelectItem key={status} value={status}>
                        {status || 'Unknown'}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>

            <div className="bg-white dark:bg-slate-900 rounded-lg border shadow-sm overflow-hidden">
              <Table>
                <TableHeader className="bg-slate-50/50 dark:bg-slate-950/50">
                  <TableRow className="hover:bg-transparent">
                    <TableHead className="w-[120px] font-semibold text-slate-900 dark:text-slate-300">NCT ID</TableHead>
                    <TableHead className="min-w-[200px] font-semibold text-slate-900 dark:text-slate-300">Study Title</TableHead>
                    <TableHead className="min-w-[150px] font-semibold text-slate-900 dark:text-slate-300">Sponsor</TableHead>
                    <TableHead className="w-[140px] font-semibold text-slate-900 dark:text-slate-300">Status</TableHead>
                    <TableHead className="w-[120px] font-semibold text-slate-900 dark:text-slate-300">Last Update</TableHead>
                    <TableHead className="min-w-[150px] font-semibold text-slate-900 dark:text-slate-300">Contact</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {filteredStudies.length === 0 ? (
                    <TableRow>
                      <TableCell colSpan={6} className="h-48 text-center text-slate-500">
                        {isFetching ? 'Loading initial batch...' : 'No trials match your search.'}
                      </TableCell>
                    </TableRow>
                  ) : (
                    filteredStudies.map((study) => (
                      <TableRow key={study.nctId} className={study.error ? "bg-red-50/30 dark:bg-red-950/10" : ""}>
                        <TableCell className="font-mono text-xs">
                          {study.error ? (
                            <span className="text-slate-500">{study.nctId}</span>
                          ) : (
                            <a 
                              href={`https://clinicaltrials.gov/study/${study.nctId}`}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="text-primary hover:text-primary/80 font-medium inline-flex items-center gap-1 hover:underline underline-offset-4"
                            >
                              {study.nctId}
                              <ExternalLink className="h-3 w-3 opacity-50" />
                            </a>
                          )}
                        </TableCell>
                        <TableCell>
                          {study.error ? (
                            <span className="text-destructive text-sm flex items-center gap-1">
                              <AlertCircle className="h-4 w-4" /> Error loading data
                            </span>
                          ) : (
                            <Tooltip>
                              <TooltipTrigger asChild>
                                <div className="line-clamp-2 text-sm text-slate-700 dark:text-slate-300 max-w-[300px] cursor-default">
                                  {study.briefTitle || <span className="text-slate-400 italic">No title</span>}
                                </div>
                              </TooltipTrigger>
                              <TooltipContent side="bottom" className="max-w-md p-3 leading-relaxed">
                                <p>{study.briefTitle}</p>
                              </TooltipContent>
                            </Tooltip>
                          )}
                        </TableCell>
                        <TableCell className="text-sm text-slate-600 dark:text-slate-400">
                          {study.leadSponsor || '-'}
                        </TableCell>
                        <TableCell>
                          <Badge variant="outline" className={`font-medium border shadow-none ${getStatusColor(study.overallStatus)}`}>
                            {study.overallStatus || 'Unknown'}
                          </Badge>
                        </TableCell>
                        <TableCell className="text-sm text-slate-600 dark:text-slate-400 tabular-nums">
                          {study.lastUpdatePostDate || '-'}
                        </TableCell>
                        <TableCell>
                          {!study.error && (study.contactName || study.contactEmail || study.contactPhone) ? (
                            <div className="space-y-1.5">
                              {study.contactName && (
                                <div className="text-sm font-medium text-slate-900 dark:text-slate-100 flex items-center gap-1.5">
                                  {study.contactName}
                                </div>
                              )}
                              {study.contactEmail && (
                                <div className="text-xs">
                                  <a href={`mailto:${study.contactEmail}`} className="text-primary hover:underline flex items-center gap-1.5">
                                    <Mail className="h-3 w-3 opacity-70" />
                                    <span className="truncate max-w-[180px]">{study.contactEmail}</span>
                                  </a>
                                </div>
                              )}
                              {study.contactPhone && (
                                <div className="text-xs text-slate-600 dark:text-slate-400 flex items-center gap-1.5">
                                  <Phone className="h-3 w-3 opacity-70" />
                                  {study.contactPhone}
                                </div>
                              )}
                            </div>
                          ) : (
                            <span className="text-sm text-slate-400 italic">-</span>
                          )}
                        </TableCell>
                      </TableRow>
                    ))
                  )}
                </TableBody>
              </Table>
            </div>
            
            {!isFetching && filteredStudies.length > 0 && (
              <div className="text-center text-sm text-slate-500 py-4">
                Showing {filteredStudies.length} {filteredStudies.length === 1 ? 'study' : 'studies'}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
