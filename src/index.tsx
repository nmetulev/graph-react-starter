import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';
import App from './App';
import * as serviceWorker from './serviceWorker';
import { Providers } from '@microsoft/mgt-element';
import { MsalProvider } from '@microsoft/mgt-msal-provider';

Providers.globalProvider = new MsalProvider({
  clientId: 'a974dfa0-9f57-49b9-95db-90f04ce2111a',
  scopes: ["User.Read", "People.Read", "Contacts.Read", "User.ReadBasic.All", "Mail.Read"]
});

ReactDOM.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
  document.getElementById('root')
);

// If you want your app to work offline and load faster, you can change
// unregister() to register() below. Note this comes with some pitfalls.
// Learn more about service workers: https://bit.ly/CRA-PWA
serviceWorker.unregister();
