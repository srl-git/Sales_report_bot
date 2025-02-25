function jsonPayload(text, url) {
  var payload = {
          "cardsV2": [
            {
              "cardId": "unique-card-id",
              "card": {
                "header": {
                  "title": "Top 15 UK Sales Last Week",
                },
                "sections": [
                  {
                    
                    "collapsible": false,
                    "uncollapsibleWidgetsCount": 1,
                    "widgets": [
                      {
                        "textParagraph": {
                          "text": text
                        }
                      },
                      {
                        "buttonList": {
                          "buttons": [
                            {
                              "text": "View full report",
                              "icon": {
                                "knownIcon": "DESCRIPTION"
                              },
                              "type": "OUTLINED",
                              "onClick": {
                                "openLink": {
                                  "url": url
                                }
                              }
                            }
                          ]
                        }
                      }
                    ]
                  }
                ]
              }
            }
          ]
      }
  
  return payload
  
}
