import React from 'react';
import './App.css';
import { Login, Person } from '@microsoft/mgt-react';
import { PersonViewType} from '@microsoft/mgt';
import { useGet, useIsSignedIn } from './mgt';
import { DetailsList, IColumn, SelectionMode, Spinner, SpinnerSize } from '@fluentui/react';
import { Message, User } from '@microsoft/microsoft-graph-types'

function App() {

  const [isSignedIn] = useIsSignedIn();
  const [user] = useGet<User>('/me');

  return (
    <div className="App">
      <header>
        <Login></Login>
      </header>
      {isSignedIn &&
        <div>
          <h3> Hello {user?.displayName},</h3>
          <h5> Here are your messages! </h5>
          <Mail></Mail>
        </div>
      }
    </div>
  );
}

function Mail() {

  let [messages, messagesLoading] = useGet('/me/messages');

  if (messagesLoading) {
    return <Spinner size={SpinnerSize.large} label="loading messages"></Spinner>
  }

  if (messages && messages.value && messages.value.length) {
    const items = messages.value.map((m: Message) => {
      return {
        key: m.id,
        from: m.sender?.emailAddress?.address,
        subject: m.subject,
        preview: m.bodyPreview
      }
    })

    const columns: IColumn[] = [
      {
        key: 'column1',
        name: 'From',
        minWidth: 150,
        maxWidth: 160,
        onRender: (item) => <Person personQuery={item.from} 
          view={PersonViewType.oneline} 
          fetchImage>
        </Person>
      },
      {
        key: 'column2',
        name: 'Subject',
        minWidth: 100,
        maxWidth: 200,
        fieldName: 'subject' 
      },
      {
        key: 'column3',
        name: 'Body',
        minWidth: 100,
        fieldName: 'preview' 
      }
    ] 

    return <div className="MessagesMain">
        <DetailsList selectionMode={SelectionMode.none} items={items} columns={columns} ></DetailsList>
      </div>
  }

  return <div>No messages</div>;
}

export default App;
