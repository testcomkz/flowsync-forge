import { createBrowserRouter, RouterProvider } from "react-router-dom";
import { AuthProvider } from './contexts/AuthContext';
import { SharePointProvider } from './contexts/SharePointContext';
import { Toaster } from './components/ui/toaster';
import { MainDashboard } from "./components/dashboard/MainDashboard";
import TubingForm from "./pages/TubingForm";
import WOForm from "./pages/WOForm";
import EditRecords from "./pages/EditRecords";
import SharePointViewer from "./pages/SharePointViewer";
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
    path: "/edit-records",
    element: <EditRecords />,
  },
  {
    path: "/sharepoint-viewer",
    element: <SharePointViewer />,
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
