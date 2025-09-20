import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { StatusBadge } from "@/components/ui/status-badge";
import { BoardItem } from "@/types";
import { Edit, Search, Filter } from "lucide-react";
import { useState } from "react";

interface BoardTableProps {
  data: BoardItem[];
  onEdit: (item: BoardItem) => void;
}

export function BoardTable({ data, onEdit }: BoardTableProps) {
  const [searchTerm, setSearchTerm] = useState("");
  const [clientFilter, setClientFilter] = useState("");
  const [rackFilter, setRackFilter] = useState("");

  const filteredData = data.filter((item) => {
    const matchesSearch = 
      item.batch.toLowerCase().includes(searchTerm.toLowerCase()) ||
      item.wo.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesClient = !clientFilter || item.client === clientFilter;
    const matchesRack = !rackFilter || item.rack === rackFilter;
    
    return matchesSearch && matchesClient && matchesRack;
  });

  const getStatusVariant = (status: string): "open" | "received" | "inspection" | "done" | "awaiting" => {
    if (status.includes("RECEIVED")) return "received";
    if (status.includes("INSPECTION DONE")) return "inspection";
    if (status.includes("AWAITING")) return "awaiting";
    if (status.includes("DONE") || status.includes("COMPLETED")) return "done";
    return "open";
  };

  const uniqueClients = [...new Set(data.map(item => item.client))];
  const uniqueRacks = [...new Set(data.map(item => item.rack))];

  return (
    <div className="space-y-4">
      {/* Filters */}
      <div className="flex flex-wrap gap-4 items-center">
        <div className="relative flex-1 max-w-sm">
          <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-muted-foreground h-4 w-4" />
          <Input
            placeholder="Search by Batch or WO..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="pl-10"
          />
        </div>
        
        <select
          value={clientFilter}
          onChange={(e) => setClientFilter(e.target.value)}
          className="px-3 py-2 border border-input rounded-md bg-background text-sm"
        >
          <option value="">All Clients</option>
          {uniqueClients.map(client => (
            <option key={client} value={client}>{client}</option>
          ))}
        </select>

        <select
          value={rackFilter}
          onChange={(e) => setRackFilter(e.target.value)}
          className="px-3 py-2 border border-input rounded-md bg-background text-sm"
        >
          <option value="">All Racks</option>
          {uniqueRacks.map(rack => (
            <option key={rack} value={rack}>{rack}</option>
          ))}
        </select>

        <Button variant="outline" size="sm">
          <Filter className="h-4 w-4 mr-2" />
          More Filters
        </Button>
      </div>

      {/* Table */}
      <div className="border rounded-lg">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Rack</TableHead>
              <TableHead>Client</TableHead>
              <TableHead>Batch</TableHead>
              <TableHead>WO</TableHead>
              <TableHead className="text-right">Qty</TableHead>
              <TableHead>Status</TableHead>
              <TableHead>Date</TableHead>
              <TableHead className="text-right">Actions</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {filteredData.length === 0 ? (
              <TableRow>
                <TableCell colSpan={8} className="text-center py-8 text-muted-foreground">
                  No data found
                </TableCell>
              </TableRow>
            ) : (
              filteredData.map((item) => (
                <TableRow key={item.id} className="hover:bg-muted/50">
                  <TableCell className="font-medium">{item.rack}</TableCell>
                  <TableCell>{item.client}</TableCell>
                  <TableCell>{item.batch}</TableCell>
                  <TableCell>{item.wo}</TableCell>
                  <TableCell className="text-right">{item.qty}</TableCell>
                  <TableCell>
                    <StatusBadge status={getStatusVariant(item.status)} />
                  </TableCell>
                  <TableCell className="text-muted-foreground">
                    {new Date(item.registered_at).toLocaleDateString()}
                  </TableCell>
                  <TableCell className="text-right">
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={() => onEdit(item)}
                    >
                      <Edit className="h-4 w-4" />
                    </Button>
                  </TableCell>
                </TableRow>
              ))
            )}
          </TableBody>
        </Table>
      </div>
      
      <div className="text-sm text-muted-foreground">
        Showing {filteredData.length} of {data.length} records
      </div>
    </div>
  );
}