import { createClient } from '@supabase/supabase-js';
import logger from '../utils/Logger';

// Initialize Supabase client
const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY;

if (!supabaseUrl || !supabaseKey) {
  throw new Error('Missing Supabase environment variables');
}

const supabase = createClient(supabaseUrl, supabaseKey);

export interface ProcessedImage {
  url: string;
  path: string;
}

class SupabaseService {
  private readonly BUCKET_NAME = 'question-images';
  public readonly storage = supabase.storage;

  async processImage(imageData: string): Promise<ProcessedImage | null> {
    try {
      let buffer: Buffer;
      let imageType: string;

      if (imageData.startsWith('data:image')) {
        // Handle base64 image data
        const matches = imageData.match(/^data:image\/([A-Za-z-+\/]+);base64,(.+)$/);
        if (!matches || matches.length !== 3) {
          throw new Error('Invalid image format');
        }
        imageType = matches[1];
        buffer = Buffer.from(matches[2], 'base64');
      } else if (imageData.startsWith('blob:')) {
        // Handle blob URLs
        try {
          const response = await fetch(imageData);
          const blob = await response.blob();
          const arrayBuffer = await blob.arrayBuffer();
          buffer = Buffer.from(arrayBuffer);
          imageType = blob.type.split('/')[1] || 'png';
        } catch (error) {
          logger.error('Error processing blob URL:', error);
          throw new Error('Failed to process blob URL');
        }
      } else {
        throw new Error('Unsupported image format');
      }

      // Generate unique filename
      const filename = `${Date.now()}-${Math.random().toString(36).substring(2)}.${imageType}`;
      const path = `${this.BUCKET_NAME}/${filename}`;

      logger.info('Uploading image to Supabase:', { 
        filename,
        contentType: `image/${imageType}`
      });

      // Upload to Supabase Storage
      const { data, error } = await this.storage
        .from(this.BUCKET_NAME)
        .upload(filename, buffer, {
          contentType: `image/${imageType}`,
          cacheControl: '3600',
          upsert: false
        });

      if (error) {
        logger.error('Error uploading to Supabase:', error);
        throw error;
      }

      // Get public URL
      const { data: { publicUrl } } = this.storage
        .from(this.BUCKET_NAME)
        .getPublicUrl(filename);

      logger.info('Successfully uploaded image:', { 
        path,
        publicUrl 
      });

      return {
        url: publicUrl,
        path: path
      };

    } catch (error) {
      logger.error('Error processing image:', error);
      return null;
    }
  }

  async deleteImage(path: string): Promise<boolean> {
    try {
      const filename = path.split('/').pop();
      if (!filename) {
        throw new Error('Invalid image path');
      }

      const { error } = await supabase.storage
        .from(this.BUCKET_NAME)
        .remove([filename]);

      if (error) {
        throw error;
      }

      return true;
    } catch (error) {
      logger.error('Error deleting image:', error);
      return false;
    }
  }
}

export const supabaseService = new SupabaseService(); 