import { useEffect, useState } from 'react'
import ComponentTable from './Components/ComponentTable'
import { ToastContainer, toast } from 'react-toastify';

function App() {
  const [user, setUser] = useState()

  useEffect(() => {
    ZOHO.CRM.CONFIG.getCurrentUser().then(function (data) {
      console.log(data);
      setUser(data)
    });
  }, [])

  // console.log({ user })
  return (
    <>
      <ToastContainer />
      <ComponentTable users={user} />
    </>
  )
}

export default App
