// user.model.ts
import { Document, Schema, model, Types } from "mongoose";

export interface IUser extends Document {
  _id: Types.ObjectId;
  email: string;
  password: string;
  phoneNumbers: string[];
  firstName: string;
  lastName: string;
  role: string;
  isActive?: boolean;
  profileImage?: string;
  signatureImage?: string;
  address?: string;
  createdBy?: Types.ObjectId;
  createdAt?: Date;
  updatedAt?: Date;
}

const userSchema = new Schema<IUser>(
  {
    email: { type: String, required: true, unique: true },
    password: { type: String, required: true },
    phoneNumbers: { type: [String], required: true },
    firstName: { type: String, required: true },
    lastName: { type: String, required: true },
    role: {
      type: String,
      required: true,
      enum: ["super_admin", "admin", "engineer", "finance", "driver"],
    },
    isActive: { type: Boolean, default: true },
    profileImage: { type: String }, // Stores S3 URL for profile image
    signatureImage: { type: String }, // Stores S3 URL for signature image
    address: { type: String },
    createdBy: { type: Schema.Types.ObjectId, ref: "User" },
  },
  { timestamps: true }
);

export const User = model<IUser>("User", userSchema);
