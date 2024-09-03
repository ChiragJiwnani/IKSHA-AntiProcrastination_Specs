const express = require('express');
const axios = require('axios');
const cors = require('cors');
const { MongoClient, ServerApiVersion } = require('mongodb');
const fs = require('fs');

const app = express();
app.use(cors());
const port = 3000; // You can change the port as needed

// Connection URI and database configuration
const uri = "mongodb+srv://ipdguys:Arh4mjJjJRKZmAeR@cluster0.ewytcbs.mongodb.net/?retryWrites=true&w=majority"; // Your MongoDB connection URI goes here
const dbName = 'AISPECS';
const collectionName = 'Tasks';

// Create a MongoClient with a MongoClientOptions object to set the Stable API version
const client = new MongoClient(uri, {
  serverApi: {
    version: ServerApiVersion.v1,
    strict: true,
    deprecationErrors: true,
  }
});

// Load JEE.json data
const jeeData = JSON.parse(fs.readFileSync('jee.json', 'utf8'));

// Function to convert time string to total seconds
function timeToSeconds(timeStr) {
  const [hours, minutes, seconds] = timeStr.split(' ').map(Number);
  return hours * 3600 + minutes * 60 + seconds;
}

// Global variable to hold the latest data
let latestData = {
  totalDifferences: {
    'No Detections': 0,
    'Book': 0,
    'Instagram': 0
  },
  recommendation: {
    remainingStudyTime: 0,
    remainingEntertainmentTime: 0
  }
};

// // Function to fetch data from MongoDB
async function fetchData() {
  try {
    await client.connect();

    const db = client.db(dbName);
    const collection = db.collection(collectionName);

    // Perform a find operation on the collection
    const tasks = await collection.find({}).toArray();

    let totalDifferences = {
      'No Detections': 0,
      'Book': 0,
      'Instagram': 0
    };

    for (let i = 0; i < tasks.length - 1; i++) {
      const timeDiff = Math.abs(timeToSeconds(tasks[i].Time) - timeToSeconds(tasks[i + 1].Time));

      // Add the time difference to the current (previous) task
      if (tasks[i].Date === tasks[i + 1].Date) {
        const taskType = tasks[i].Task;
        totalDifferences[taskType] += timeDiff;
      }
    }

    // Calculate the time spent on study and entertainment
    const studyTime = totalDifferences['Book'];
    const entertainmentTime = totalDifferences['Instagram'];
    console.log(entertainmentTime)
    console.log("hello")
    // Calculate the remaining time based on jee.json targets
    const remainingStudyTime = jeeData.study - studyTime;
    const remainingEntertainmentTime = jeeData.entertainment - entertainmentTime;
    console.log(remainingEntertainmentTime)
    const recommendation = {
      remainingStudyTime: remainingStudyTime > 0 ? remainingStudyTime : 0,
      remainingEntertainmentTime: remainingEntertainmentTime > 0 ? remainingEntertainmentTime : 0
    };

    // Update the global variable with the latest data
    latestData = {
      totalDifferences,
      recommendation
    };
  } catch (error) {
    console.error(error);
  } finally {
    await client.close();
  }
}
//////////////////////////////////////////


// async function fetchData() {
//   try {
//     await client.connect();

//     const db = client.db(dbName);
//     const collection = db.collection(collectionName);

//     // Get the current date
//     const currentDate = new Date();
//     const currentDateString = currentDate.toLocaleDateString('en-US', {
//       day: '2-digit',
//       month: '2-digit',
//       year: '2-digit'
//     });

//     // Construct a query to find documents with the Date field matching the current date
//     const query = { Date: currentDateString };

//     // Find documents matching the query
//     const tasks = await collection.find(query).toArray();

//     let totalDifferences = {
//       'No Detections': 0,
//       'Book': 0,
//       'Instagram': 0
//     };

//     for (let i = 0; i < tasks.length - 1; i++) {
//       const timeDiff = Math.abs(timeToSeconds(tasks[i].Time) - timeToSeconds(tasks[i + 1].Time));

//       // Add the time difference to the current (previous) task
//       if (tasks[i].Date === tasks[i + 1].Date) {
//         const taskType = tasks[i].Task;
//         totalDifferences[taskType] += timeDiff;
//       }
//     }

//     // Calculate the time spent on study and entertainment
//     const studyTime = totalDifferences['Book'];
//     const entertainmentTime = totalDifferences['Instagram'];

//     // Calculate the remaining time based on jee.json targets
//     const remainingStudyTime = jeeData.study - studyTime;
//     const remainingEntertainmentTime = jeeData.entertainment - entertainmentTime;

//     const recommendation = {
//       remainingStudyTime: Math.max(remainingStudyTime, 0),
//       remainingEntertainmentTime: Math.max(remainingEntertainmentTime, 0)
//     };

//     // Update the global variable with the latest data
//     latestData = {
//       totalDifferences,
//       recommendation
//     };
//   } catch (error) {
//     console.error(error);
//   } finally {
//     await client.close();
//   }
// }


app.get('/api/tasks/instadata', async (req, res) => {
  try {
    await client.connect();

    const db = client.db(dbName);
    const collection = db.collection(collectionName);

    // Fetch Instagram data
    const instaData = await collection.find({ Task: 'Instagram' }).toArray();

    // Fetch Book data
    const bookData = await collection.find({ Task: 'Book' }).toArray();

    // Close the MongoDB connection
    await client.close();

    // Send both instaData and bookData in the response
    res.send({ instaData, bookData });
  } catch (error) {
    console.error(error);
    res.status(500).send('Internal Server Error');
  }
});





// Initial fetch and setup interval to fetch data every 5 seconds
fetchData();
setInterval(fetchData, 5000);

// Endpoint for fetching data and providing recommendations
app.get('/api/tasks', (req, res) => {
  res.send(latestData);
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
