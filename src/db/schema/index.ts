import { pgEnum, pgTable, text, timestamp, uuid, integer, boolean } from 'drizzle-orm/pg-core';
import { relations, sql } from 'drizzle-orm';

export const userRoleEnum = pgEnum('role', ['user', 'admin']);
export const membershipEnum = pgEnum('membership', ['free', 'pro']);

export const profilesTable = pgTable('profiles', {
  userId: text('user_id').primaryKey().notNull(),
  firstName: text('first_name').notNull(),
  lastName: text('last_name').notNull(),
  email: text('email').notNull().unique(),
  role: userRoleEnum('role').default('user').notNull(),
  membership: membershipEnum('membership').default('free').notNull(),
  createdAt: timestamp('created_at').defaultNow().notNull(),
  updatedAt: timestamp('updated_at').defaultNow().notNull()
});

export const questionsTable = pgTable("questions", {
  id: uuid("id").defaultRandom().primaryKey(),
  userId: text("user_id")
    .references(() => profilesTable.userId, { onDelete: "cascade" })
    .notNull(),
  categoryId: uuid("category_id").default(sql`null`),
  title: text("title").notNull(),
  body: text("body").notNull(),
  images: text("images"),
  votes: integer("votes").default(0).notNull(),
  createdAt: timestamp("created_at").defaultNow().notNull(),
  updatedAt: timestamp("updated_at")
    .defaultNow()
    .notNull()
    .$onUpdate(() => new Date())
});

export const answersTable = pgTable("answers", {
  id: uuid("id").defaultRandom().primaryKey(),
  questionId: uuid("question_id")
    .references(() => questionsTable.id, { onDelete: "cascade" })
    .notNull(),
  userId: text("user_id")
    .references(() => profilesTable.userId, { onDelete: "cascade" })
    .notNull(),
  body: text("body").notNull(),
  images: text("images"),
  votes: integer("votes").default(0).notNull(),
  accepted: boolean("accepted").default(false).notNull(),
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

export type Profile = typeof profilesTable.$inferSelect;
export type InsertProfile = typeof profilesTable.$inferInsert;
export type Question = typeof questionsTable.$inferSelect;
export type InsertQuestion = typeof questionsTable.$inferInsert;
export type Answer = typeof answersTable.$inferSelect;
export type InsertAnswer = typeof answersTable.$inferInsert; 