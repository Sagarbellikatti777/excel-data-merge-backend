const express = require('express');
const cors = require('cors');
const mergeRoutes = require('./routes/merge');

const app = express();

// CORS Options Configuration
const corsOptions = {
  origin: 'https://marge-excel-data-withoutsaving-user.vercel.app', // Allow only your frontend domain
  methods: ['GET', 'POST'], // Allow only GET and POST requests (adjust if needed)
  allowedHeaders: ['Content-Type', 'Authorization'], // Allow Content-Type and Authorization headers
  credentials: true, // Allow credentials like cookies (optional, depends on your use case)
};

// Enable CORS with the specified options
app.use(cors(corsOptions));

// Middleware to parse incoming JSON requests
app.use(express.json());

// Route for merging Excel files
app.use('/api', mergeRoutes);

// Start the server
const PORT = process.env.PORT || 5000; // Use environment port or default to 5000
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
