const express = require('express');
const cors = require('cors');
const mergeRoutes = require('./routes/merge');

const app = express();

// Allow requests from Vercel frontend
const corsOptions = {
  origin: 'https://marge-excel-data-withoutsaving-user.vercel.app', // Frontend URL (replace with your Vercel URL)
  methods: ['GET', 'POST'], // Allowed HTTP methods
  allowedHeaders: ['Content-Type', 'Authorization'], // Allowed headers
  credentials: true, // Allow credentials (optional, depends on your setup)
};

app.use(cors(corsOptions)); // Use CORS with the specified options
app.use(express.json());

app.use('/api', mergeRoutes);

const PORT = process.env.PORT || 5000; // Ensure this matches your server's PORT
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
