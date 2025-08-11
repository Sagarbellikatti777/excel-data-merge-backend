const express = require('express');
const cors = require('cors');
const mergeRoutes = require('./routes/merge');

const app = express();
app.use(cors({
  origin: '*', // Allows requests from any domain
}));
app.use(express.json());

app.use('/api', mergeRoutes);

const PORT = 5000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
