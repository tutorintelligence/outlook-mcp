/**
 * Improved search emails functionality
 */
const config = require('../config');
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');

/**
 * Search emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchEmails(args) {
  const folder = args.folder || "inbox";
  const requestedCount = args.count || 10;
  const query = args.query || '';
  const from = args.from || '';
  const to = args.to || '';
  const subject = args.subject || '';
  const hasAttachments = args.hasAttachments;
  const unreadOnly = args.unreadOnly;
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Resolve the folder path
    const endpoint = await resolveFolderPath(accessToken, folder);
    console.error(`Using endpoint: ${endpoint} for folder: ${folder}`);
    
    // Execute progressive search with pagination
    const response = await progressiveSearch(
      endpoint, 
      accessToken, 
      { query, from, to, subject },
      { hasAttachments, unreadOnly },
      requestedCount
    );
    
    return formatSearchResults(response);
  } catch (error) {
    // Handle authentication errors
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    // General error response
    return {
      content: [{ 
        type: "text", 
        text: `Error searching emails: ${error.message}`
      }]
    };
  }
}

/**
 * Execute a search with progressively simpler fallback strategies
 * @param {string} endpoint - API endpoint
 * @param {string} accessToken - Access token
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 * @param {number} maxCount - Maximum number of results to retrieve
 * @returns {Promise<object>} - Search results
 */
async function progressiveSearch(endpoint, accessToken, searchTerms, filterTerms, maxCount) {
  const hasSearchTerms = searchTerms.query || searchTerms.from || searchTerms.to || searchTerms.subject;
  const hasBooleanFilters = filterTerms.hasAttachments === true || filterTerms.unreadOnly === true;

  // If no search criteria provided, just return recent emails
  if (!hasSearchTerms && !hasBooleanFilters) {
    console.error("No search criteria provided, returning recent emails");
    const basicParams = {
      $top: Math.min(50, maxCount),
      $select: config.EMAIL_SELECT_FIELDS,
      $orderby: 'receivedDateTime desc'
    };
    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, basicParams, maxCount);
    response._searchInfo = { strategy: 'recent-emails', reason: 'no-criteria' };
    return response;
  }

  // 1. Try $search with all terms combined (most specific)
  // NOTE: $search and $orderby cannot be used together on the messages endpoint
  if (hasSearchTerms) {
    try {
      const params = buildSearchParams(searchTerms, filterTerms, Math.min(50, maxCount));
      console.error("Attempting combined search with params:", JSON.stringify(params));

      const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, params, maxCount);
      console.error(`Combined search returned ${response.value?.length || 0} results`);
      response._searchInfo = { strategy: 'combined-search' };
      return response;
    } catch (error) {
      console.error(`Combined search failed: ${error.message}`);
    }

    // 2. Try each search term individually
    const searchPriority = ['from', 'to', 'subject', 'query'];
    for (const term of searchPriority) {
      if (searchTerms[term]) {
        try {
          console.error(`Attempting search with only ${term}: "${searchTerms[term]}"`);
          const simplifiedParams = {
            $top: Math.min(50, maxCount),
            $select: config.EMAIL_SELECT_FIELDS
            // No $orderby — incompatible with $search
          };

          if (term === 'query') {
            simplifiedParams.$search = `"${searchTerms[term]}"`;
          } else {
            simplifiedParams.$search = `${term}:"${searchTerms[term]}"`;
          }

          addBooleanFilters(simplifiedParams, filterTerms);

          const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, simplifiedParams, maxCount);
          console.error(`Search with ${term} returned ${response.value?.length || 0} results`);
          response._searchInfo = { strategy: `single-term-${term}` };
          return response;
        } catch (error) {
          console.error(`Search with ${term} failed: ${error.message}`);
        }
      }
    }
  }

  // 3. Try with only boolean filters (these use $filter, compatible with $orderby)
  if (hasBooleanFilters) {
    try {
      console.error("Attempting search with only boolean filters");
      const filterOnlyParams = {
        $top: Math.min(50, maxCount),
        $select: config.EMAIL_SELECT_FIELDS,
        $orderby: 'receivedDateTime desc'
      };
      addBooleanFilters(filterOnlyParams, filterTerms);

      const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, filterOnlyParams, maxCount);
      console.error(`Boolean filter search returned ${response.value?.length || 0} results`);
      response._searchInfo = { strategy: 'boolean-filters-only' };
      return response;
    } catch (error) {
      console.error(`Boolean filter search failed: ${error.message}`);
    }
  }

  // 4. All strategies threw errors — fall back to recent emails
  console.error("All search strategies failed with errors, falling back to recent emails");
  const basicParams = {
    $top: Math.min(50, maxCount),
    $select: config.EMAIL_SELECT_FIELDS,
    $orderby: 'receivedDateTime desc'
  };
  const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, basicParams, maxCount);
  response._searchInfo = { strategy: 'recent-emails', reason: 'all-strategies-errored' };
  return response;
}

/**
 * Build search parameters from search terms and filter terms
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 * @param {number} count - Maximum number of results
 * @returns {object} - Query parameters
 */
function buildSearchParams(searchTerms, filterTerms, count) {
  const params = {
    $top: count,
    $select: config.EMAIL_SELECT_FIELDS
  };

  // Build KQL search terms
  const kqlTerms = [];

  if (searchTerms.query) {
    kqlTerms.push(`"${searchTerms.query}"`);
  }

  if (searchTerms.subject) {
    kqlTerms.push(`subject:"${searchTerms.subject}"`);
  }

  if (searchTerms.from) {
    kqlTerms.push(`from:"${searchTerms.from}"`);
  }

  if (searchTerms.to) {
    kqlTerms.push(`to:"${searchTerms.to}"`);
  }

  if (kqlTerms.length > 0) {
    // $search and $orderby cannot be used together on messages endpoint
    params.$search = kqlTerms.join(' ');
  } else {
    // No search terms — safe to use $orderby
    params.$orderby = 'receivedDateTime desc';
  }

  // Add boolean filters
  addBooleanFilters(params, filterTerms);

  return params;
}

/**
 * Add boolean filters to query parameters
 * @param {object} params - Query parameters
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 */
function addBooleanFilters(params, filterTerms) {
  const filterConditions = [];
  
  if (filterTerms.hasAttachments === true) {
    filterConditions.push('hasAttachments eq true');
  }
  
  if (filterTerms.unreadOnly === true) {
    filterConditions.push('isRead eq false');
  }
  
  // Add $filter parameter if we have any filter conditions
  if (filterConditions.length > 0) {
    params.$filter = filterConditions.join(' and ');
  }
}

/**
 * Format search results into a readable text format
 * @param {object} response - The API response object
 * @returns {object} - MCP response object
 */
function formatSearchResults(response) {
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ 
        type: "text", 
        text: `No emails found matching your search criteria.`
      }]
    };
  }
  
  // Format results
  const emailList = response.value.map((email, index) => {
    const sender = email.from?.emailAddress || { name: 'Unknown', address: 'unknown' };
    const date = new Date(email.receivedDateTime).toLocaleString();
    const readStatus = email.isRead ? '' : '[UNREAD] ';
    
    return `${index + 1}. ${readStatus}${date} - From: ${sender.name} (${sender.address})\nSubject: ${email.subject}\nID: ${email.id}\n`;
  }).join("\n");
  
  // Add search strategy info if available
  let additionalInfo = '';
  if (response._searchInfo) {
    additionalInfo = `\n(Search used ${response._searchInfo.strategy} strategy)`;
  }
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} emails matching your search criteria:${additionalInfo}\n\n${emailList}`
    }]
  };
}

module.exports = handleSearchEmails;
