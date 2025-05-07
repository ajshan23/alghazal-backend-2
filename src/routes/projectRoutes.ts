import express from "express";
import {
  createProject,
  getProjects,
  getProject,
  updateProject,
  updateProjectStatus,
  updateProjectProgress,
  deleteProject,
  assignProject,
  getEngineerProjects,
  getProjectProgressUpdates,
  generateInvoiceData,
} from "../controllers/projectController";
import { authenticate, authorize } from "../middlewares/authMiddleware";

const router = express.Router();

// Apply authentication to all routes
router.use(authenticate);

// Create project - Admin/Engineer only
router.post(
  "/",
  authorize(["admin", "super_admin", "engineer"]),
  createProject
);

// Get all projects
router.get("/", getProjects);
router.get("/engineer", getEngineerProjects);

// Get single project
router.get("/:id", getProject);
router.get(
  "/:projectId/invoice",
  authorize(["admin", "super_admin", "finance", "engineer"]),
  generateInvoiceData
);
// Update project - Admin/Engineer only
router.put(
  "/:id",
  authorize(["admin", "super_admin", "engineer"]),
  updateProject
);
router.post(
  "/:id/assign",
  authorize(["admin", "super_admin", "finance"]),
  assignProject
);
router.get(
  "/:projectId/progress",
  authorize(["admin", "super_admin", "finance"]),
  getProjectProgressUpdates
);
// Update project status
router.patch(
  "/:id/status",
  authorize(["admin", "super_admin", "engineer", "finance"]),
  updateProjectStatus
);

// Update project progress
router.patch(
  "/:id/progress",
  authorize(["admin", "super_admin", "engineer"]),
  updateProjectProgress
);

// Delete project - Admin only
router.delete("/:id", authorize(["admin", "super_admin"]), deleteProject);

export default router;
