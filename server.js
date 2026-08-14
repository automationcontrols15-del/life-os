const express = require('express');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Serve life-os.html as the root page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'life-os.html'));
});

app.listen(PORT, () => {
  console.log(`Life OS running on port ${PORT}`);
});
