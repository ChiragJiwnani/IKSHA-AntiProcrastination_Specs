import { React, useEffect, useState } from 'react';
import { Bar, Line } from 'react-chartjs-2';
import 'chart.js/auto';
import axios from 'axios';

export default function ChartCompt() {
  const fixUrl = "http://192.168.31.204";
  const getUrl = `${fixUrl}:3000/api/tasks`;
  const getInstaUrl = `${fixUrl}:3000/api/tasks/instadata`

  const [instadata, setInstaData] = useState([]);
  const [bookdata, setBookData] = useState([]);
  const [remainingStudyTime, setRemainingStudyTime] = useState(0);
  const [remainingEntertainmentTime, setRemainingEntertainmentTime] = useState(0);
  const [newInstaData , setNewInstaData] = useState(0)
  const [newBookData , setNewBookData] = useState(0)
  const fetchData = async () => {
    try {
      const response = await axios.get(getUrl);
      const data = response.data;
      console.log(data.totalDifferences['Instagram']);

      setInstaData(data.totalDifferences['Instagram']);
      setBookData(data.totalDifferences['Book']);
      setRemainingStudyTime(data.recommendation['remainingStudyTime']);
      setRemainingEntertainmentTime(data.recommendation['remainingEntertainmentTime']);
    } catch (error) {
      console.log(error);
    }
  };

  useEffect(() => {
    fetchData();
    console.log('Data fetched');
    console.log(instadata, bookdata, remainingStudyTime, remainingEntertainmentTime)
  }, []);

  const fetchInstaData = async () => {
    try {
      const response = await axios.get(getInstaUrl);
      const data = response.data;
      if (data) {
        console.log("Received data from backend:", data);
        // Set both instaData and bookData in the component's state
        setNewInstaData(data.instaData);
        setNewBookData(data.bookData);
      }
    } catch (error) {
      console.log("Error fetching data:", error);
    }
  };
  
  useEffect(() => {
    fetchInstaData();
  }, []);
  



  let dailyAvg = (instadata + bookdata) / 60;
  let weeklyTotal = dailyAvg * 7;
  const instaTimeSpentOnDays  = [0, 0, 0, 0, 0, 0, 0]; // Mon-Sun
  if (Array.isArray(newInstaData)) {
    // Initialize an array to hold the time spent on each day
   
    // Aggregate the time spent on each day
    newInstaData.forEach(item => {
        const dateParts = item.Date.split(' '); // Split the date string by space
        const day = parseInt(dateParts[0]); // Extract day
        const month = parseInt(dateParts[1]) -7; // Extract month (months are zero-based in JavaScript)
        const year = parseInt(dateParts[2]); // Extract year

        // Create a new Date object using extracted day, month, and year
        const date = new Date(day, month, year);

        // Get the day of the week (0 for Sunday, 1 for Monday, etc.)
        const dayOfWeek = date.getDay();

        // Aggregate the time spent on this day
        instaTimeSpentOnDays [dayOfWeek] += parseInt(item.Time);
    })}

    const bookTimeSpentOnDays = [0, 0, 0, 0, 0, 0, 0]; // Mon-Sun

    // Aggregate the time spent on each day for bookData
    if (Array.isArray(newBookData)) {
        newBookData.forEach(item => {
            const dateParts = item.Date.split(' '); // Split the date string by space
            const day = parseInt(dateParts[0]); // Extract day
            const month = parseInt(dateParts[1]) - 7; // Extract month (months are zero-based in JavaScript)
            const year = parseInt(dateParts[2]); // Extract year
    
            // Create a new Date object using extracted day, month, and year
            const date = new Date(day, month, year);
    
            // Get the day of the week (0 for Sunday, 1 for Monday, etc.)
            const dayOfWeek = date.getDay();
    
            // Aggregate the time spent on this day
            bookTimeSpentOnDays[dayOfWeek] += parseInt(item.Time);
        });
    }  

  return (
    <>
      <div>
        <Bar
          data={{
            labels: ['Mon', 'Tues', 'Wed', 'Thur' ,'Fri' ,'Sat' , 'Sun'],
            datasets: [
              {
                label: "Time spent on Mobile (hrs)",
                data: instaTimeSpentOnDays 
              }
            ]
          }}
          height={400}
          width={600}
          className='chart-container'
        />
      </div>
      <div>
        <Bar
          data={{
            labels: ['Mon', 'Tues', 'Wed', 'Thur', 'Fri', 'Sat', 'Sun'],
            datasets: [
              {
                label: 'Instagram',
                data: instaTimeSpentOnDays 
              },
              {
                label: 'Books',
                data: bookTimeSpentOnDays
              }
            ]
          }}
          height={400}
          width={600}
          className='line-container'
        />
      </div>
      {/* <div>
        <Bar
          data={{
            labels: ['Mon', 'Tues', 'Wed', 'Thur', 'Fri', 'Sat', 'Sun'],
            datasets: [
              {
                label: 'Remaining Study Time (hrs)',
                data: [remainingStudyTime , remainingStudyTime , remainingStudyTime , remainingStudyTime , remainingStudyTime , remainingStudyTime , remainingStudyTime ]
              },
              {
                label: 'Remaining Entertainment Time (hrs)',
                data: [remainingEntertainmentTime , remainingEntertainmentTime , remainingEntertainmentTime , remainingEntertainmentTime , remainingEntertainmentTime , remainingEntertainmentTime , remainingEntertainmentTime ]
              }
            ]
          }}
          height={400}
          width={600}
          className='chart-container'
        />
      </div> */}

      <div className='stats'>
        <div>
          <p>Daily Avg</p>
          <h4>{dailyAvg.toFixed(2)}<sub>m</sub></h4>
        </div>
        <div>
          <p>Weekly Total</p>
          <h4>{Math.floor(weeklyTotal)}<sub>H</sub> {Math.floor((weeklyTotal % 1) * 60)}<sub>m</sub></h4>
        </div>
        <div>
          <p>The Shortest</p>
          <h4>20<sub>H</sub> 39<sub>m</sub></h4>
        </div>
        <div>
          <p>The Longest</p>
          <h4>40<sub>H</sub> 39<sub>m</sub></h4>
        </div>
      </div>
    </>
  );
}
