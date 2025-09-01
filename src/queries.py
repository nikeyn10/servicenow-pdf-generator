# GraphQL query strings and builders

GET_STATUS_COLUMN = '''
query GetStatusColumn($boardId: [ID!]) {
  boards(ids: $boardId) {
    columns(ids: "status95") { id title type settings_str }
  }
}
'''

GET_ITEMS_PAGE = '''
query GetItemsPage($boardId:[ID!], $limit:Int!) {
  boards(ids:$boardId) {
    items_page(limit:$limit) {
      cursor
      items {
        id
        name
        assets { id name file_extension public_url url }
        column_values(ids:["date_mkt2sps1","date_mktr60pn","status95"]) {
          id
          text
          ... on StatusValue { index label }
        }
      }
    }
  }
}
'''

NEXT_ITEMS_PAGE = '''
query NextItems($cursor:String!, $limit:Int!){
  next_items_page(cursor:$cursor, limit:$limit){
    cursor
    items {
      id
      name
      assets { id name file_extension public_url url }
      column_values(ids:["date_mkt2sps1","date_mktr60pn","status95"]) {
        id
        text
        ... on StatusValue { index label }
      }
    }
  }
}
'''
