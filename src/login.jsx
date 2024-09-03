import React, { useState } from 'react'


const Login = ({onLogin})=> {
const [usertext,setUserText] = useState('')
const [passtext,setPassText] = useState('')
const [failtext,setFailText] = useState(true)

const handleUserInput = (e) =>{
setUserText(e.target.value)
console.log(e.target.value)
}
const handlePassInput = (e) =>{
  setPassText(e.target.value)
  console.log(e.target.value)
}
const clear = () =>{
  console.log("clicked!")
setUserText("")
setPassText("")

}

const handleLogin =() =>{
  console.log("clicked!")
   if(usertext==="ipd" && passtext==="dhruv"){
    console.log(usertext)
    onLogin(true)
  }
  else{
     onLogin(false)
     setFailText(false)

   }
}

    return (
    <div className='login-page'>
      <div className='login-container'>
      <p className='login'>Login</p>
      <div className='credentials'>
        <div className='userid'>User ID:
        <input type='text' onChange={(e)=>{handleUserInput(e)}} value={usertext}/>
        </div>
        <div className='pass'>Password:
        <input type='password' onChange={(e)=>{handlePassInput(e)}} value={passtext}/>

        <button type='submit' className='btn-sub' onClick={handleLogin}>
        Submit    
        </button>
        <button type='submit' className='btn-cl' onClick={clear}>
        Clear    
        </button>
        </div>
      </div>
      </div>
      {failtext?null:<h7>Please Try Again</h7>}
    </div>
  )
}

export default Login