import { pgEnum, pgTable, text, timestamp } from 'drizzle-orm/pg-core';

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

export type Profile = typeof profilesTable.$inferSelect;
export type InsertProfile = typeof profilesTable.$inferInsert; 