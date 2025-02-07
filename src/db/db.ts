import { drizzle } from 'drizzle-orm/node-postgres';
import { Pool } from 'pg';
import * as schema from './schema';
import logger from '../utils/Logger';

// Initialize the connection pool
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: {
    rejectUnauthorized: false // Required for self-signed certs
  }
});

// Test the connection
pool.connect()
  .then(() => {
    logger.info('Successfully connected to Supabase Postgres database');
  })
  .catch((error) => {
    logger.error('Failed to connect to database:', error);
  });

// Create the database instance
export const db = drizzle(pool, { schema });

// Export the pool for potential direct usage
export { pool }; 