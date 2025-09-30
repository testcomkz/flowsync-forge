I am building a React application connected to Microsoft Azure (SharePoint/Excel through Microsoft Graph API).  
Currently, I must click a button to fetch the latest data. I want to remove this button and make my app update data in **real time**, but without overloading the server by polling every second.  

Please:  
1. Implement **real-time updates** for my data using the most efficient Azure-supported solution.  
   - Prefer **Microsoft Graph API webhooks** (subscription to Excel/SharePoint changes).  
   - Or suggest a reliable alternative (SignalR, WebSockets, Azure Event Grid, etc.) if better for my case.  
2. Show how to configure the webhook subscription so my backend receives change notifications when data in SharePoint Excel is updated.  
3. Add a lightweight backend (Node.js/Express or Next.js API route) that listens to the webhook notifications.  
4. Push changes from the backend to the React frontend via **WebSocket/SignalR**, so the UI updates automatically only when data actually changes.  
5. Provide a full working example (backend + React frontend) with code snippets.  
6. Make sure the solution is **scalable** and **does not overload** the database or server.  
