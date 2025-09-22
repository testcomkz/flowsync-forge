import { useState } from "react";
import { Header } from "@/components/layout/Header";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { ArrowLeft, Search } from "lucide-react";
import { useNavigate } from "react-router-dom";

// Mock data for existing records
const mockRecords = [
  { id: 1, client: "Dunga", wo_no: "2245", batch: "Batch # 1", diameter: "3 1/2", qty: 120 },
  { id: 2, client: "Dunga", wo_no: "2245", batch: "Batch # 2", diameter: "3 1/2", qty: 120 },
  { id: 3, client: "KenSary", wo_no: "2200", batch: "Batch # 1", diameter: "2 7/8", qty: 158 },
  { id: 4, client: "Tasbulat", wo_no: "2200", batch: "Batch # 74", diameter: "2 7/8", qty: 158 }
];

export default function EditRecords() {
  const navigate = useNavigate();
  const [searchFilters, setSearchFilters] = useState({
    client: "",
    wo_no: "",
    batch: ""
  });
  const [selectedRecord, setSelectedRecord] = useState<any>(null);
  const [filteredRecords, setFilteredRecords] = useState(mockRecords);

  const handleSearch = () => {
    let filtered = mockRecords;
    
    if (searchFilters.client) {
      filtered = filtered.filter(record => 
        record.client.toLowerCase().includes(searchFilters.client.toLowerCase())
      );
    }
    
    if (searchFilters.wo_no) {
      filtered = filtered.filter(record => 
        record.wo_no.includes(searchFilters.wo_no)
      );
    }
    
    if (searchFilters.batch) {
      filtered = filtered.filter(record => 
        record.batch.toLowerCase().includes(searchFilters.batch.toLowerCase())
      );
    }
    
    setFilteredRecords(filtered);
  };

  const handleEdit = (record: any) => {
    setSelectedRecord({ ...record });
  };

  const handleSave = () => {
    // TODO: Implement SharePoint update
    console.log("Updating record:", selectedRecord);
    setSelectedRecord(null);
    // Refresh the list
    handleSearch();
  };

  const handleFilterChange = (field: string, value: string) => {
    setSearchFilters(prev => ({ ...prev, [field]: value }));
  };

  const handleRecordChange = (field: string, value: string) => {
    setSelectedRecord((prev: any) => ({ ...prev, [field]: value }));
  };

  return (
    <div className="min-h-screen bg-gray-50">
      <Header />
      <div className="container mx-auto px-6 py-8">
        <div className="mb-6">
          <Button 
            variant="outline" 
            onClick={() => navigate("/")}
            className="flex items-center space-x-2"
          >
            <ArrowLeft className="w-4 h-4" />
            <span>Back to Dashboard</span>
          </Button>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Search and List */}
          <Card>
            <CardHeader>
              <CardTitle className="text-xl font-bold">Search Records</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              {/* Search Filters */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className="space-y-2">
                  <Label htmlFor="search_client">Client</Label>
                  <Select onValueChange={(value) => handleFilterChange("client", value === "__ALL__" ? "" : value)}>
                    <SelectTrigger>
                      <SelectValue placeholder="All clients" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="__ALL__">All clients</SelectItem>
                      <SelectItem value="Dunga">Dunga</SelectItem>
                      <SelectItem value="KenSary">KenSary</SelectItem>
                      <SelectItem value="Tasbulat">Tasbulat</SelectItem>
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="search_wo">Work Order</Label>
                  <Input
                    id="search_wo"
                    value={searchFilters.wo_no}
                    onChange={(e) => handleFilterChange("wo_no", e.target.value)}
                    placeholder="Enter WO number"
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="search_batch">Batch</Label>
                  <Input
                    id="search_batch"
                    value={searchFilters.batch}
                    onChange={(e) => handleFilterChange("batch", e.target.value)}
                    placeholder="Enter batch"
                  />
                </div>
              </div>

              <Button onClick={handleSearch} className="w-full flex items-center space-x-2">
                <Search className="w-4 h-4" />
                <span>Search Records</span>
              </Button>

              {/* Results List */}
              <div className="space-y-2 max-h-96 overflow-y-auto">
                {filteredRecords.map((record) => (
                  <div 
                    key={record.id}
                    className="p-3 border rounded-lg hover:bg-gray-50 cursor-pointer"
                    onClick={() => handleEdit(record)}
                  >
                    <div className="font-medium">{record.client} - {record.wo_no}</div>
                    <div className="text-sm text-gray-600">
                      {record.batch} | {record.diameter} | Qty: {record.qty}
                    </div>
                  </div>
                ))}
              </div>
            </CardContent>
          </Card>

          {/* Edit Form */}
          <Card>
            <CardHeader>
              <CardTitle className="text-xl font-bold">
                {selectedRecord ? "Edit Record" : "Select a Record to Edit"}
              </CardTitle>
            </CardHeader>
            <CardContent>
              {selectedRecord ? (
                <div className="space-y-4">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="space-y-2">
                      <Label htmlFor="edit_client">Client</Label>
                      <Select 
                        value={selectedRecord.client}
                        onValueChange={(value) => handleRecordChange("client", value)}
                      >
                        <SelectTrigger>
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="Dunga">Dunga</SelectItem>
                          <SelectItem value="KenSary">KenSary</SelectItem>
                          <SelectItem value="Tasbulat">Tasbulat</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>

                    <div className="space-y-2">
                      <Label htmlFor="edit_wo">Work Order</Label>
                      <Input
                        id="edit_wo"
                        value={selectedRecord.wo_no}
                        onChange={(e) => handleRecordChange("wo_no", e.target.value)}
                      />
                    </div>

                    <div className="space-y-2">
                      <Label htmlFor="edit_batch">Batch</Label>
                      <Input
                        id="edit_batch"
                        value={selectedRecord.batch}
                        onChange={(e) => handleRecordChange("batch", e.target.value)}
                      />
                    </div>

                    <div className="space-y-2">
                      <Label htmlFor="edit_diameter">Diameter</Label>
                      <Select 
                        value={selectedRecord.diameter}
                        onValueChange={(value) => handleRecordChange("diameter", value)}
                      >
                        <SelectTrigger>
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="3 1/2">3 1/2"</SelectItem>
                          <SelectItem value="2 7/8">2 7/8"</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>

                    <div className="space-y-2">
                      <Label htmlFor="edit_qty">Quantity</Label>
                      <Input
                        id="edit_qty"
                        type="number"
                        value={selectedRecord.qty}
                        onChange={(e) => handleRecordChange("qty", e.target.value)}
                      />
                    </div>
                  </div>

                  <div className="flex justify-end space-x-4 pt-4">
                    <Button 
                      type="button" 
                      variant="outline" 
                      onClick={() => setSelectedRecord(null)}
                    >
                      Cancel
                    </Button>
                    <Button onClick={handleSave}>
                      Save Changes
                    </Button>
                  </div>
                </div>
              ) : (
                <div className="text-center text-gray-500 py-8">
                  <p>Select a record from the list to edit its details</p>
                </div>
              )}
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  );
}
