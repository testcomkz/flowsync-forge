import { Header } from "@/components/layout/Header";
import { MainDashboard } from "@/components/dashboard/MainDashboard";

export default function Index() {
  return (
    <div className="min-h-screen bg-gray-50">
      <Header />
      <MainDashboard />
    </div>
  );
}