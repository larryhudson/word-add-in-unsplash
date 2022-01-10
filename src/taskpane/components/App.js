import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton, TextField } from "@fluentui/react";

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      searchTerm: '',
      searchResults: [],
      loading: false,
    };
  }

  insertImage = async (image) => {
 
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */
      this.setState({loading: true})

      // download image and convert data to base64
      const getBase64FromUrl = async (url) => {
        const data = await fetch(url);
        const blob = await data.blob();
        return new Promise((resolve) => {
          const reader = new FileReader();
          reader.readAsDataURL(blob); 
          reader.onloadend = () => {
            const base64data = reader.result;   
            resolve(base64data.replace(/data:.+?,/, ''));
          }
        });
      }

      const imageBase64Data = await getBase64FromUrl(image.urls.raw)
      const insertedImage = context.document.getSelection().insertInlinePictureFromBase64(imageBase64Data, 'After');
      insertedImage.altTextDescription = image.alt_description
      this.setState({loading: false})
      await context.sync();
      
    });
  }

  search = async () => {
    this.setState({loading: true})
    const unsplashApiKey = `YOUR_UNSPLASH_API_KEY_GOES_HERE`
    const response = await fetch(`https://api.unsplash.com/search/photos?page=1&query=${this.state.searchTerm}`, {
      headers: {
        Authorization: `Client-ID ${unsplashApiKey}`
      }
    }).then(r => r.json())
    this.setState({
      searchResults: response.results,
      loading: false,
    })
  }

  // 
  render() {
    const { title } = this.props;

    return (
      <div>
          <label htmlFor="search-input">Search term</label>
          <TextField id="search-input" type="text" onChange={(e) => this.setState({searchTerm: e.target.value})}  />
          <DefaultButton iconProps={{iconName: "Search" }} onClick={this.search}>Search</DefaultButton>
          {this.state.loading && (<p>Loading...</p>)}
          {this.state.searchResults && (
            <div style={{display: 'grid', gridTemplateColumns: 'repeat(auto-fit, 200px)'}}>
              {this.state.searchResults.map(searchResult => (
                <button onClick={() => this.insertImage(searchResult)}>
                  <img src={searchResult.urls.thumb} alt={searchResult.alt_description} />
                </button>
              ))}
            </div>
          )}
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
