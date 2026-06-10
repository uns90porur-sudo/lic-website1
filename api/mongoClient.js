const { MongoClient } = require('mongodb');

// Get the connection string from Vercel Environment Variables or fallback to hardcoded string
const uri = process.env.MONGODB_URI || "mongodb+srv://admin:LIC123@cluster0.ljzkkqf.mongodb.net/?appName=Cluster0";
const options = {};

let client;
let clientPromise;

if (!uri) {
  console.error('Please add your Mongo URI to Vercel Environment Variables');
} else {
  if (process.env.NODE_ENV === 'development') {
    if (!global._mongoClientPromise) {
      client = new MongoClient(uri, options);
      global._mongoClientPromise = client.connect();
    }
    clientPromise = global._mongoClientPromise;
  } else {
    client = new MongoClient(uri, options);
    clientPromise = client.connect();
  }
}

module.exports = clientPromise;
