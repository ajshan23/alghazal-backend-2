import express from "express";
import {
  createClient,
  getClients,
  getClient,
  updateClient,
  deleteClient,
  getClientByTrn,
  getClientsByPincode,
} from "../controllers/clientController";
import { authenticate, authorize } from "../middlewares/authMiddleware";

const router = express.Router();

// Apply authentication to all routes
router.use(authenticate);

// Create client - Admin only
router.post("/", authorize(["admin", "super_admin"]), createClient);

// Get all clients (with optional search and pincode filter)
router.get("/", getClients);

// Get client by ID
router.get("/:id", getClient);

// Get client by TRN
router.get("/trn/:trnNumber", getClientByTrn);

// Get clients by pincode
router.get("/pincode/:pincode", getClientsByPincode);

// Update client - Admin only
router.put("/:id", authorize(["admin", "super_admin"]), updateClient);

// Delete client - Admin only
router.delete("/:id", authorize(["admin", "super_admin"]), deleteClient);

export default router;
