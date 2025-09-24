import { createBrowserRouter, RouterProvider, Navigate } from "react-router-dom";
import { AuthProvider } from './contexts/AuthContext';
import { SharePointProvider } from './contexts/SharePointContext';
import { Toaster } from './components/ui/toaster';
import { MainDashboard } from "./components/dashboard/MainDashboard";
import TubingForm from "./pages/TubingForm";
import WOForm from "./pages/WOForm";
import EditRecords from "./pages/EditRecords";
import SharePointViewer from "./pages/SharePointViewer";
import InspectionData from "./pages/InspectionData";
import LoadOut from "./pages/LoadOut";
import LoadOutEdit from "./pages/LoadOutEdit";
import InspectionEdit from "./pages/InspectionEdit";
import TubingRegistryEdit from "./pages/TubingRegistryEdit";
import "./index.css";

const router = createBrowserRouter([
  {
    path: "/",
    element: <MainDashboard />,
  },
  {
    path: "/tubing-form",
    element: <TubingForm />,
  },
  {
    path: "/wo-form",
    element: <WOForm />,
  },
  {
    path: "/edit",
    element: <Navigate to="/edit-records" replace />,
  },
  {
    path: "/edit-records",
    element: <EditRecords />,
  },
  {
    path: "/sharepoint-viewer",
    element: <SharePointViewer />,
  },
  {
    path: "/inspection-data",
    element: <InspectionData />,
  },
  {
    path: "/load-out",
    element: <LoadOut />,
  },
  {
    path: "/load-out-edit",
    element: <LoadOutEdit />,
  },
  {
    path: "/inspection-edit",
    element: <InspectionEdit />,
  },
  {
    path: "/tubing-registry-edit",
    element: <TubingRegistryEdit />,
  },
]);

function App() {
  return (
    <AuthProvider>
      <SharePointProvider>
        <RouterProvider router={router} />
        <Toaster />
      </SharePointProvider>
    </AuthProvider>
  );
}

export default App;
