import { pgTable, uuid, text, integer, timestamp, boolean } from 'drizzle-orm/pg-core';
import { profilesTable } from './profiles-schema';
import { questionsTable } from './questions-schema';
import { relations } from 'drizzle-orm';

export const answersTable = pgTable("answers", {
  id: uuid("id").defaultRandom().primaryKey(),
  questionId: uuid("question_id")
    .references(() => questionsTable.id, { onDelete: "cascade" })
    .notNull(),
  userId: text("user_id")
    .references(() => profilesTable.userId, { onDelete: "cascade" })
    .notNull(),
  body: text("body").notNull(),
  images: text("images"), // Keeping this as is since it exists in the table
  votes: integer("votes").default(0).notNull(),
  accepted: boolean("accepted").default(false).notNull(),
  createdAt: timestamp("created_at").defaultNow().notNull(),
  updatedAt: timestamp("updated_at")
    .defaultNow()
    .notNull()
    .$onUpdate(() => new Date())
});

export const answersRelations = relations(answersTable, ({ one }) => ({
  question: one(questionsTable, {
    fields: [answersTable.questionId],
    references: [questionsTable.id],
  }),
  profile: one(profilesTable, {
    fields: [answersTable.userId],
    references: [profilesTable.userId],
  }),
}));

export type Answer = typeof answersTable.$inferSelect;
export type InsertAnswer = typeof answersTable.$inferInsert; 