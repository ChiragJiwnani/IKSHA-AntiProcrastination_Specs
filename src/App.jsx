import { useState } from 'react'
import reactLogo from './assets/react.svg'
import ChartComp from './chart.jsx'
import Login from './login.jsx'
import viteLogo from '/vite.svg'

import './App.css'

function App() {
const [isAuth,setIsAuth] = useState(true)
const handleLogin = (status) =>{
  setIsAuth(status)
}
  return (
    <>
      
      {isAuth?(

        <div className="main-container">
        <div className="header">
          <h5>Ikhsha</h5>
        </div>
        <div className="dashboard-container">
          <div className="dashboard"><h4>Dashboard</h4></div>
       <ChartComp /> 

        </div>
      </div>
      ):(
        <Login onLogin={handleLogin}/>
      )}
     
    </>
  )
}

export default App
