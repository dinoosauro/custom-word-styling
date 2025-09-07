import { mount } from 'svelte'
import './app.css'
import App from './App.svelte'
import Presentation from './Presentation.svelte'

const app = mount(new URLSearchParams(window.location.search).get("ForceOfficeUI") ? App : Presentation, {
  target: document.getElementById('app')!,
})

export default app
