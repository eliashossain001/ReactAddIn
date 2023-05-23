import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props: AppProps) {
    super(props);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [],
    });
  }

  click = async () => {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph(
        "Taskpane",
        Word.InsertLocation.end
      );

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  };

  render() {
    return (
      <div className="ms-welcome">
        <Header
          logo="assets/logo-filled.png"
          title={this.props.title}
          message=""
        />
        <HeroList message="" items={this.state.listItems}>
          {/* Render other components here */}
        </HeroList>
      </div>
    );
  }
}
