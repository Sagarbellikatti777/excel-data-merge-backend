const express = require('express');
const cors = require('cors');
const mergeRoutes = require('./routes/merge'); // Ensure this is correct

const app = express();

// Allow requests from the Vercel frontend domain
const corsOptions = {
  origin: 'https://marge-excel-data-withoutsaving-user.vercel.app', // Frontend URL
  methods: ['GET', 'POST'], // Allowed methods
  allowedHeaders: ['Content-Type', 'Authorization'], // Allowed headers
  credentials: true, // Allow credentials (cookies, etc.) if needed
};

app.use(cors(corsOptions)); // Use CORS with the options

app.use(express.json());

// API routes
app.use('/api', mergeRoutes);

const PORT = process.env.PORT || 5000; // Ensure this matches your server's PORT
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
