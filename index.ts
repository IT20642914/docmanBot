import app from "./app";
import { startDocumentApiServer } from "./services/documentApiServer";

// Start the application
(async () => {
  startDocumentApiServer(app);
  await app.start();
  console.log(`\nBot started, app listening to`, process.env.PORT || process.env.port || 3978);
})();
