import { pgTable, uuid, text, integer, timestamp } from 'drizzle-orm/pg-core';
import { profilesTable } from './profiles-schema';
import { sql, relations } from "drizzle-orm"
import { answersTable } from './answers-schema';

export const questionsTable = pgTable("questions", {
  id: uuid("id").defaultRandom().primaryKey(),
  userId: text("user_id")
    .references(() => profilesTable.userId, { onDelete: "cascade" })
    .notNull(),
  categoryId: uuid("category_id").default(sql`null`),
  title: text("title").notNull(),
  body: text("body").notNull(),
  images: text("images"), // Keeping this as is since it exists in the table
  votes: integer("votes").default(0).notNull(),
  createdAt: timestamp("created_at").defaultNow().notNull(),
  updatedAt: timestamp("updated_at")
    .defaultNow()
    .notNull()
    .$onUpdate(() => new Date())
});

export const questionsRelations = relations(questionsTable, ({ many, one }) => ({
  answers: many(answersTable),
  profile: one(profilesTable, {
    fields: [questionsTable.userId],
    references: [profilesTable.userId],
  }),
}));

export type Question = typeof questionsTable.$inferSelect;
export type InsertQuestion = typeof questionsTable.$inferInsert; 