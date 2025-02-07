import express, { Request, Response } from 'express';
import multer from 'multer';
import { supabaseService } from '../services/supabase-service';
import logger from '../utils/Logger';

const router = express.Router();

// Configure multer for memory storage
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 5 * 1024 * 1024 // 5MB limit
  },
  fileFilter: (req, file, cb) => {
    // Accept only image files
    if (!file.mimetype.startsWith('image/')) {
      return cb(new Error('Only image files are allowed!'));
    }
    cb(null, true);
  }
});

// Handle image upload
router.post('/', upload.single('image'), async (req: Request, res: Response) => {
  try {
    const file = (req as any).file;
    if (!file) {
      return res.status(400).json({ error: 'No image file provided' });
    }

    // Convert buffer to base64
    const base64Image = `data:${file.mimetype};base64,${file.buffer.toString('base64')}`;

    // Upload to Supabase
    const result = await supabaseService.processImage(base64Image);
    
    if (!result) {
      throw new Error('Failed to upload image to Supabase');
    }

    logger.info('Image uploaded to Supabase successfully', { 
      url: result.url,
      path: result.path 
    });

    return res.status(200).json({ url: result.url });
  } catch (error) {
    logger.error('Error uploading file:', error);
    return res.status(500).json({ error: 'Error uploading file' });
  }
});

export default router;
