import 'office-ui-fabric-react/dist/css/fabric.min.css';
import App from './components/App';
import { AppContainer } from 'react-hot-loader';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import * as React from 'react';
import * as ReactDOM from 'react-dom';

initializeIcons();

let isOfficeInitialized = false;

const title = 'Phishing Net for Office 365';


const render = (Component) => {
    ReactDOM.render(
        <AppContainer>
            <Component title={title} isOfficeInitialized={isOfficeInitialized} />
        </AppContainer>,
        document.getElementById('container')
    );
};

console.log("Office is now being initialized!");

/* Render application after Office initializes */
(async () => {
    await Office.onReady();
    isOfficeInitialized = true;
    console.log("Office is ready!")
    render(App);
})();

//Office.initialize = () => {
//    isOfficeInitialized = true;
//    render(App);
//};

/* Initial render showing a progress bar */
render(App); 

if ((module as any).hot) {
    (module as any).hot.accept('./components/App', () => {
        const NextApp = require('./components/App').default;
        render(NextApp);
    });
}
