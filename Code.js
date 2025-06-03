// Google Apps Script - Code.gs
// Account Audit Document Generator - Complete Version

// Test function to use simplified HTML
function doGetSimple() {
  try {
    return HtmlService.createHtmlOutputFromFile('Index_Simple')
      .setTitle('Account Audit Generator')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Error in doGetSimple:', error);
    return HtmlService.createHtmlOutput('<h1>Error loading application</h1><p>' + error.toString() + '</p>');
  }
}

// Test deployment with basic HTML
function doGetTest() {
  return HtmlService.createHtmlOutputFromFile('Test')
    .setTitle('Test Page')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// Test basic HTML loading
function testHtmlLoading() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Index');
    console.log('Index.html loaded successfully');
    return true;
  } catch (e) {
    console.error('Error loading Index.html:', e);
    return false;
  }
}

function doGet() {
  try {
    // Create template and evaluate to process includes
    const template = HtmlService.createTemplateFromFile('Index_Styled');
    return template.evaluate()
      .setTitle('Account Audit Generator')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Error in doGet:', error);
    return HtmlService.createHtmlOutput('<h1>Error loading application</h1><p>' + error.toString() + '</p>');
  }
}

// Include function for HTML templates
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Original version with Material UI (may have compatibility issues)
function doGetOriginal() {
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Account Audit Generator')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Error in doGetOriginal:', error);
    return HtmlService.createHtmlOutput('<h1>Error loading application</h1><p>' + error.toString() + '</p>');
  }
}

function processAuditData(jsonData) {
  try {
    // Parse the JSON data
    const data = JSON.parse(jsonData);
    
    // Template configuration
    const templateId = '1Ij4vM7EBqTxJ-TOiHNiWlRUiC-Hf2S8NU4Gull_jWH0';
    
    // Check if we have access to DriveApp
    if (!DriveApp) {
      throw new Error('DriveApp is not available. Please ensure the script has the necessary permissions.');
    }
    
    // Try to access the template with better error handling
    let templateFile;
    try {
      templateFile = DriveApp.getFileById(templateId);
    } catch (e) {
      console.error('Error accessing template:', e);
      throw new Error('Cannot access template document. Please check: 1) The template ID is correct, 2) You have access to the template, 3) The script has Drive permissions.');
    }
    
    // Create a copy of the template
    const siteName = data.account_overview?.site || 'Unknown';
    const timestamp = new Date().toLocaleDateString().replace(/\//g, '-');
    const newFileName = `Account Audit - ${siteName} - ${timestamp}`;
    
    let newFile;
    try {
      newFile = templateFile.makeCopy(newFileName);
    } catch (e) {
      console.error('Error creating copy:', e);
      throw new Error('Cannot create a copy of the template. Please check Drive permissions.');
    }
    
    // Open the new document
    let doc;
    try {
      doc = DocumentApp.openById(newFile.getId());
    } catch (e) {
      console.error('Error opening document:', e);
      throw new Error('Cannot open the new document. Please check DocumentApp permissions.');
    }
    
    const body = doc.getBody();
    
    // Process all replacements
    const { textReplacements, tableReplacements } = processDocumentReplacements(data);
    
    // Apply text replacements first
    let replacementCount = 0;
    for (const [placeholder, value] of Object.entries(textReplacements)) {
      try {
        body.replaceText(placeholder, value);
        replacementCount++;
      } catch (e) {
        console.error(`Error replacing ${placeholder}:`, e);
      }
    }
    
    // Apply table replacements
    for (const [placeholder, tableData] of Object.entries(tableReplacements)) {
      try {
        replaceTextWithTable(body, placeholder, tableData);
        replacementCount++;
      } catch (e) {
        console.error(`Error replacing table ${placeholder}:`, e);
      }
    }
    
    console.log(`Successfully replaced ${replacementCount} placeholders`);
    
    // Save and close
    doc.saveAndClose();
    
    // Get the URL of the new document
    const docUrl = doc.getUrl();
    
    // Return success message with link
    return `Document created successfully! <a href="${docUrl}" target="_blank">Open Document</a>`;
    
  } catch (error) {
    console.error('Error in processAuditData:', error);
    throw new Error('Failed to process data: ' + error.toString());
  }
}

function replaceTextWithTable(body, searchText, tableData) {
  // Find all occurrences of the placeholder
  var foundElement = body.findText(searchText);
  
  while (foundElement) {
    var element = foundElement.getElement();
    var parent = element.getParent();
    
    // Get the index of the parent in the body
    var parentIndex = body.getChildIndex(parent);
    
    // Clear the placeholder text
    element.asText().setText('');
    
    // Insert the table after the parent element
    if (tableData && tableData.length > 0) {
      var table = body.insertTable(parentIndex + 1, tableData);
      
      // Style the table
      table.setBorderWidth(1);
      table.setBorderColor('#000000');
      
      // Style header row if exists
      if (tableData.length > 0) {
        var headerRow = table.getRow(0);
        for (var i = 0; i < headerRow.getNumCells(); i++) {
          headerRow.getCell(i).setBackgroundColor('#f0f0f0');
          headerRow.getCell(i).getChild(0).asParagraph().setBold(true);
        }
      }
    }
    
    // Remove the empty paragraph if it's completely empty
    if (parent.asText().getText().trim() === '') {
      body.removeChild(parent);
    }
    
    // Search for next occurrence
    foundElement = body.findText(searchText, foundElement);
  }
}

function processDocumentReplacements(data) {
  const textReplacements = {};
  const tableReplacements = {};
  
  // Helper function to safely get nested values
  function safeGet(obj, path, defaultValue = 'Not configured') {
    try {
      const keys = path.split('.');
      let result = obj;
      
      for (const key of keys) {
        if (result === null || result === undefined) {
          return defaultValue;
        }
        result = result[key];
      }
      
      if (result === null || result === undefined) {
        return defaultValue;
      }
      
      return result;
    } catch (e) {
      return defaultValue;
    }
  }
  
  // Helper function to format values for display
  function formatValue(value, defaultValue = 'Not configured') {
    if (value === null || value === undefined) {
      return defaultValue;
    }
    
    if (typeof value === 'boolean') {
      return value ? 'Yes' : 'No';
    }
    
    if (typeof value === 'number') {
      return value.toString();
    }
    
    if (Array.isArray(value)) {
      if (value.length === 0) return 'None';
      return value.map(v => formatValue(v)).join(', ');
    }
    
    if (typeof value === 'object') {
      // Handle special object cases
      if (value.hasOwnProperty('configured')) {
        return value.configured ? 'Configured' : 'Not configured';
      }
      return JSON.stringify(value, null, 2);
    }
    
    return String(value);
  }
  
  // 1. Account Overview - ALL fields
  textReplacements['{{site}}'] = formatValue(safeGet(data, 'account_overview.site'));
  textReplacements['{{account_id}}'] = formatValue(safeGet(data, 'account_overview.account_id'));
  textReplacements['{{account_mode}}'] = formatValue(safeGet(data, 'account_overview.account_mode'));
  textReplacements['{{billing_account_id}}'] = formatValue(safeGet(data, 'account_overview.billing_account_id'));
  textReplacements['{{created_at}}'] = formatValue(safeGet(data, 'account_overview.created_at'));
  textReplacements['{{timezone}}'] = formatValue(safeGet(data, 'account_overview.timezone'));
  textReplacements['{{currency_type}}'] = formatValue(safeGet(data, 'account_overview.currency_type'));
  
  // 2. User Management - NEW SECTION
  textReplacements['{{total_users}}'] = formatValue(safeGet(data, 'user_management.total_users'));
  textReplacements['{{active_users_count}}'] = formatValue(safeGet(data, 'user_management.active_users_count'));
  textReplacements['{{inactive_users_count}}'] = formatValue(safeGet(data, 'user_management.inactive_users_count'));
  
  // Users by Role Table
  const usersByRole = safeGet(data, 'user_management.users_by_role', {});
  const usersByRoleData = [['Role', 'Count']];
  for (const [role, count] of Object.entries(usersByRole)) {
    usersByRoleData.push([role, formatValue(count)]);
  }
  tableReplacements['{{users_by_role_table}}'] = usersByRoleData;
  
  // User Details Table
  const userDetails = safeGet(data, 'user_management.role_details', []);
  const userDetailsData = [['Email', 'Role', 'Active', 'Last Sign In', 'Sign In Count']];
  userDetails.forEach(user => {
    userDetailsData.push([
      formatValue(user.email),
      formatValue(user.role),
      formatValue(user.is_active),
      formatValue(user.last_sign_in_at),
      formatValue(user.sign_in_count)
    ]);
  });
  tableReplacements['{{user_details_table}}'] = userDetailsData;
  
  // 3. Campaign Configuration - EXPANDED
  textReplacements['{{total_campaigns}}'] = formatValue(safeGet(data, 'campaign_configuration.total_count'));
  textReplacements['{{active_campaigns}}'] = formatValue(safeGet(data, 'campaign_configuration.active_campaigns'));
  textReplacements['{{paused_campaigns}}'] = formatValue(safeGet(data, 'campaign_configuration.paused_campaigns'));
  textReplacements['{{archived_campaigns}}'] = formatValue(safeGet(data, 'campaign_configuration.archived_campaigns'));
  textReplacements['{{holdout_campaigns}}'] = formatValue(safeGet(data, 'campaign_configuration.holdout_campaigns'));
  
  // Campaigns by Execution Term Table
  const byExecTerm = safeGet(data, 'campaign_configuration.by_exec_term', {});
  const execTermData = [['Execution Term', 'Count']];
  for (const [term, count] of Object.entries(byExecTerm)) {
    execTermData.push([term, formatValue(count)]);
  }
  tableReplacements['{{campaigns_by_exec_term_table}}'] = execTermData;
  
  // Campaigns by Status Table
  const byStatus = safeGet(data, 'campaign_configuration.by_status', {});
  const statusData = [['Status', 'Count']];
  for (const [status, count] of Object.entries(byStatus)) {
    statusData.push([status, formatValue(count)]);
  }
  tableReplacements['{{campaigns_by_status_table}}'] = statusData;
  
  // Keep existing campaign table for compatibility
  const campaignTableData = [['Campaign Type', 'Count']];
  if (data.campaign_configuration) {
    const config = data.campaign_configuration;
    campaignTableData.push(['Total Campaigns', formatValue(config.total_count || 0)]);
    campaignTableData.push(['Active', formatValue(config.active_campaigns || 0)]);
    campaignTableData.push(['Paused', formatValue(config.paused_campaigns || 0)]);
    campaignTableData.push(['Archived', formatValue(config.archived_campaigns || 0)]);
    campaignTableData.push(['Holdout', formatValue(config.holdout_campaigns || 0)]);
  }
  tableReplacements['{{campaignTable}}'] = campaignTableData;
  
  // 4. Channel Configuration - COMPLETE
  textReplacements['{{allow_channel_email}}'] = formatValue(safeGet(data, 'channel_configuration.email.enabled'));
  textReplacements['{{max_email_channel}}'] = formatValue(safeGet(data, 'channel_configuration.email.max_per_day'));
  textReplacements['{{max_email_channel_unit}}'] = formatValue(safeGet(data, 'channel_configuration.email.max_unit'));
  
  textReplacements['{{allow_channel_sms}}'] = formatValue(safeGet(data, 'channel_configuration.sms.enabled'));
  textReplacements['{{max_sms_channel}}'] = formatValue(safeGet(data, 'channel_configuration.sms.max_per_day'));
  textReplacements['{{max_sms_channel_unit}}'] = formatValue(safeGet(data, 'channel_configuration.sms.max_unit'));
  
  textReplacements['{{allow_channel_push}}'] = formatValue(safeGet(data, 'channel_configuration.push.enabled'));
  textReplacements['{{max_push_channel}}'] = formatValue(safeGet(data, 'channel_configuration.push.max_per_day'));
  textReplacements['{{max_push_channel_unit}}'] = formatValue(safeGet(data, 'channel_configuration.push.max_unit'));
  
  textReplacements['{{allow_channel_in_app}}'] = formatValue(safeGet(data, 'channel_configuration.in_app.enabled'));
  textReplacements['{{max_in_app_channel}}'] = formatValue(safeGet(data, 'channel_configuration.in_app.max_per_day'));
  textReplacements['{{max_in_app_channel_unit}}'] = formatValue(safeGet(data, 'channel_configuration.in_app.max_unit'));
  
  textReplacements['{{allow_channel_cloud_app}}'] = formatValue(safeGet(data, 'channel_configuration.display.enabled'));
  textReplacements['{{max_display_channel}}'] = formatValue(safeGet(data, 'channel_configuration.display.max_per_day'));
  textReplacements['{{max_display_channel_unit}}'] = formatValue(safeGet(data, 'channel_configuration.display.max_unit'));
  
  textReplacements['{{allow_channel_forms}}'] = formatValue(safeGet(data, 'channel_configuration.forms.enabled'));
  textReplacements['{{allow_channel_live_content}}'] = formatValue(safeGet(data, 'channel_configuration.live_content.enabled'));
  
  textReplacements['{{max_all_channels}}'] = formatValue(safeGet(data, 'channel_configuration.all_channels.max_per_period'));
  textReplacements['{{max_all_channels_unit}}'] = formatValue(safeGet(data, 'channel_configuration.all_channels.max_unit'));
  
  // 5. Data Retention - EXPANDED
  const esRetention = safeGet(data, 'data_retention.es_data_retention_in_days', {});
  textReplacements['{{es_data_retention_in_days_useractions}}'] = formatValue(esRetention.useractions);
  textReplacements['{{es_data_retention_in_days_product_events}}'] = formatValue(esRetention.product_events);
  textReplacements['{{es_data_retention_in_days_events}}'] = formatValue(esRetention.events);
  textReplacements['{{es_data_retention_in_days_raw_events}}'] = formatValue(safeGet(data, 'data_retention.es_data_retention_in_days_raw_events'));
  textReplacements['{{transactions_data_retention_period}}'] = formatValue(safeGet(data, 'data_retention.transactions_data_retention_period'));
  textReplacements['{{cookie_only_user_retention_days}}'] = formatValue(safeGet(data, 'data_retention.cookie_only_user_retention_days'));
  
  // 6. Features - COMPLETE
  textReplacements['{{allow_feature_api}}'] = formatValue(safeGet(data, 'features.allow_feature_api'));
  textReplacements['{{allow_feature_segmentation}}'] = formatValue(safeGet(data, 'features.allow_feature_segmentation'));
  textReplacements['{{allow_feature_personalization}}'] = formatValue(safeGet(data, 'features.allow_feature_personalization'));
  textReplacements['{{allow_feature_predictive}}'] = formatValue(safeGet(data, 'features.allow_feature_predictive'));
  textReplacements['{{allow_feature_promotions}}'] = formatValue(safeGet(data, 'features.allow_feature_promotions'));
  textReplacements['{{allow_feature_shared_assets}}'] = formatValue(safeGet(data, 'features.allow_feature_shared_assets'));
  textReplacements['{{allow_personalization_studio}}'] = formatValue(safeGet(data, 'features.allow_personalization_studio'));
  textReplacements['{{allow_predictive_affinities}}'] = formatValue(safeGet(data, 'features.allow_predictive_affinities'));
  textReplacements['{{allow_segment_precompute}}'] = formatValue(safeGet(data, 'features.allow_segment_precompute'));
  textReplacements['{{enable_catalogs}}'] = formatValue(safeGet(data, 'features.enable_catalogs'));
  textReplacements['{{enable_customer_groups}}'] = formatValue(safeGet(data, 'features.enable_customer_groups'));
  textReplacements['{{enable_event_verification}}'] = formatValue(safeGet(data, 'features.enable_event_verification'));
  textReplacements['{{enable_external_fetch}}'] = formatValue(safeGet(data, 'features.enable_external_fetch'));
  textReplacements['{{enable_global_holdout}}'] = formatValue(safeGet(data, 'features.enable_global_holdout'));
  textReplacements['{{enable_holdouts}}'] = formatValue(safeGet(data, 'features.enable_holdouts'));
  textReplacements['{{enable_interests}}'] = formatValue(safeGet(data, 'features.enable_interests'));
  textReplacements['{{enable_predictive_studio}}'] = formatValue(safeGet(data, 'features.enable_predictive_studio'));
  textReplacements['{{enable_recommendations_v2}}'] = formatValue(safeGet(data, 'features.enable_recommendations_v2'));
  textReplacements['{{enable_segment_prioritization}}'] = formatValue(safeGet(data, 'features.enable_segment_prioritization'));
  textReplacements['{{enable_send_time_optimization}}'] = formatValue(safeGet(data, 'features.enable_send_time_optimization'));
  textReplacements['{{enable_syndication}}'] = formatValue(safeGet(data, 'features.enable_syndication'));
  textReplacements['{{enable_anniversaries}}'] = formatValue(safeGet(data, 'features.enable_anniversaries'));
  textReplacements['{{enable_autocomplete}}'] = formatValue(safeGet(data, 'features.enable_autocomplete'));
  textReplacements['{{catalog_activity_enabled}}'] = formatValue(safeGet(data, 'features.catalog_activity_enabled'));
  textReplacements['{{anniversary_attributes_mapping}}'] = formatValue(safeGet(data, 'features.anniversary_attributes_mapping'));
  textReplacements['{{event_verification_list}}'] = formatValue(safeGet(data, 'features.event_verification_list'));
  
  // 7. Integration Settings - NEW COMPLETE SECTION
  textReplacements['{{total_integrations}}'] = formatValue(safeGet(data, 'integration_settings.total_integrations'));
  textReplacements['{{facebook_status}}'] = formatValue(safeGet(data, 'integration_settings.facebook_status'));
  textReplacements['{{google_status}}'] = formatValue(safeGet(data, 'integration_settings.google_status'));
  textReplacements['{{webhook_configured}}'] = formatValue(safeGet(data, 'integration_settings.webhook_configured'));
  
  // S3 Credentials
  const s3Creds = safeGet(data, 'integration_settings.s3_credentials', {});
  textReplacements['{{s3_configured}}'] = formatValue(s3Creds.configured);
  textReplacements['{{s3_credentials_access_key}}'] = formatValue(s3Creds.access_key || safeGet(data, 'account_configuration.s3_credentials_access_key'));
  textReplacements['{{s3_credentials_updated_at}}'] = formatValue(s3Creds.updated_at || safeGet(data, 'account_configuration.s3_credentials_updated_at'));
  textReplacements['{{s3_credentials_expires_at}}'] = formatValue(s3Creds.expires_at || safeGet(data, 'account_configuration.s3_credentials_expires_at'));
  
  // Integrations List Table
  const integrationsList = safeGet(data, 'integration_settings.integrations_list', []);
  const integrationsData = [['Integration', 'Status', 'Type']];
  integrationsList.forEach(integration => {
    integrationsData.push([
      formatValue(integration.name),
      formatValue(integration.status),
      formatValue(integration.integration_type)
    ]);
  });
  if (integrationsList.length === 0) {
    integrationsData.push(['No integrations configured', '', '']);
  }
  tableReplacements['{{integrations_list_table}}'] = integrationsData;
  
  // 8. Predictive Features
  textReplacements['{{predictive_scores_count}}'] = formatValue(safeGet(data, 'predictive_features.predictive_scores_count'));
  textReplacements['{{active_scores}}'] = formatValue(safeGet(data, 'predictive_features.active_scores'));
  
  // Predictive Score Table
  const predictiveTableData = [['Score Name', 'Type', 'Status']];
  const predictiveScores = data.predictive_features?.predictive_scores || [];
  if (predictiveScores.length > 0) {
    predictiveScores.forEach(score => {
      predictiveTableData.push([
        formatValue(score.name),
        formatValue(score.score_type),
        formatValue(score.status)
      ]);
    });
  } else {
    predictiveTableData.push(['No predictive scores configured', '', '']);
  }
  tableReplacements['{{predictiveScore}}'] = predictiveTableData;
  
  // 9. Segmentation Configuration - EXPANDED
  textReplacements['{{total_segments}}'] = formatValue(safeGet(data, 'segmentation_configuration.total_segments'));
  textReplacements['{{active_segments}}'] = formatValue(safeGet(data, 'segmentation_configuration.active_segments'));
  textReplacements['{{segments_with_precompute}}'] = formatValue(safeGet(data, 'segmentation_configuration.segments_with_precompute'));
  textReplacements['{{max_segment_reference_breadth}}'] = formatValue(safeGet(data, 'segmentation_configuration.max_segment_reference_breadth'));
  textReplacements['{{max_segment_reference_depth}}'] = formatValue(safeGet(data, 'segmentation_configuration.max_segment_reference_depth'));
  textReplacements['{{max_prioritized_segments}}'] = formatValue(safeGet(data, 'segmentation_configuration.max_prioritized_segments'));
  textReplacements['{{segmentation_event_keys_exclude_list}}'] = formatValue(safeGet(data, 'segmentation_configuration.segmentation_event_keys_exclude_list'));
  textReplacements['{{segmentation_product_attributes_exclude_list}}'] = formatValue(safeGet(data, 'segmentation_configuration.segmentation_product_attributes_exclude_list'));
  textReplacements['{{segmentation_user_attributes_exclude_list}}'] = formatValue(safeGet(data, 'segmentation_configuration.segmentation_user_attributes_exclude_list'));
  
  // 10. Event Configuration - NEW SECTION
  textReplacements['{{custom_events_count}}'] = formatValue(safeGet(data, 'event_configuration.custom_events_count'));
  textReplacements['{{goal_events_whitelist}}'] = formatValue(safeGet(data, 'event_configuration.goal_events_whitelist'));
  textReplacements['{{product_event_whitelist}}'] = formatValue(safeGet(data, 'event_configuration.product_event_whitelist'));
  textReplacements['{{archive_events}}'] = formatValue(safeGet(data, 'event_configuration.archive_events'));
  textReplacements['{{event_verification_enabled}}'] = formatValue(safeGet(data, 'event_configuration.event_verification_enabled'));
  
  // Custom Events Table
  const customEventsList = safeGet(data, 'event_configuration.custom_events_list', []);
  const customEventsData = [['Event Name', 'Event Type', 'Active']];
  customEventsList.forEach(event => {
    customEventsData.push([
      formatValue(event.name),
      formatValue(event.event_type),
      formatValue(event.is_active)
    ]);
  });
  if (customEventsList.length === 0) {
    customEventsData.push(['No custom events configured', '', '']);
  }
  tableReplacements['{{custom_events_table}}'] = customEventsData;
  
  // 11. Attribution & Analytics
  textReplacements['{{attribution_days}}'] = formatValue(safeGet(data, 'attribution_analytics.attribution_days'));
  textReplacements['{{attribution_type}}'] = formatValue(safeGet(data, 'attribution_analytics.attribution_type'));
  textReplacements['{{dashboard_max_active_users}}'] = formatValue(safeGet(data, 'attribution_analytics.dashboard_max_active_users'));
  
  // 12. Messaging Configuration
  textReplacements['{{max_triggers}}'] = formatValue(safeGet(data, 'messaging_configuration.max_triggers'));
  textReplacements['{{allow_bypassing_dedupe_check}}'] = formatValue(safeGet(data, 'messaging_configuration.allow_bypassing_dedupe_check'));
  
  // Trigger level custom URL params
  const urlParams = safeGet(data, 'messaging_configuration.trigger_level_custom_url_params', null);
  if (urlParams) {
    const paramsList = [];
    for (const [key, value] of Object.entries(urlParams)) {
      paramsList.push(`${key}: ${value.default || 'N/A'}`);
    }
    textReplacements['{{trigger_level_custom_url_params}}'] = paramsList.join(', ');
  } else {
    textReplacements['{{trigger_level_custom_url_params}}'] = 'Not configured';
  }
  
  // Message Limits Table
  const messageLimits = safeGet(data, 'messaging_configuration.messaging_limits', []);
  const messageLimitsData = [['Channel', 'Limit Type', 'Max Messages', 'Time Period']];
  messageLimits.forEach(limit => {
    if (limit.max_messages || limit.limit_type) { // Only include if there's actual data
      messageLimitsData.push([
        formatValue(limit.channel),
        formatValue(limit.limit_type),
        formatValue(limit.max_messages),
        formatValue(limit.time_period)
      ]);
    }
  });
  if (messageLimitsData.length === 1) {
    messageLimitsData.push(['No message limits configured', '', '', '']);
  }
  tableReplacements['{{message_limits_table}}'] = messageLimitsData;
  
  // 13. Recommendation System - EXPANDED
  textReplacements['{{total_recipes}}'] = formatValue(safeGet(data, 'recommendation_system.total_recipes'));
  textReplacements['{{active_recipes}}'] = formatValue(safeGet(data, 'recommendation_system.active_recipes'));
  textReplacements['{{inactive_recipes}}'] = formatValue(safeGet(data, 'recommendation_system.inactive_recipes'));
  
  // Recipes by Category Table
  const recipesByCategory = safeGet(data, 'recommendation_system.recipes_by_category', {});
  const categoryData = [['Category', 'Count']];
  for (const [category, count] of Object.entries(recipesByCategory)) {
    categoryData.push([category, formatValue(count)]);
  }
  if (Object.keys(recipesByCategory).length === 0) {
    categoryData.push(['No categories', '0']);
  }
  tableReplacements['{{recipes_by_category_table}}'] = categoryData;
  
  // Recommendation Table
  const recommendationTableData = [['Name', 'Category', 'Active']];
  const recommendations = data.recommendation_system?.recipes_detail || [];
  if (recommendations.length > 0) {
    recommendations.forEach(rec => {
      recommendationTableData.push([
        formatValue(rec.name),
        formatValue(rec.category),
        formatValue(rec.is_active)
      ]);
    });
  } else {
    recommendationTableData.push(['No recommendations configured', '', '']);
  }
  tableReplacements['{{recommendationTable}}'] = recommendationTableData;
  
  // 14. Transactions Table
  const transactionTableData = [['Description', 'Transaction Name', 'Event', 'Active']];
  const transactions = data.transactions || [];
  if (transactions.length > 0) {
    transactions.forEach(trans => {
      transactionTableData.push([
        formatValue(trans.description),
        formatValue(trans.transaction_name),
        formatValue(trans.event),
        formatValue(trans.active)
      ]);
    });
  } else {
    transactionTableData.push(['No transactions configured', '', '', '']);
  }
  tableReplacements['{{transactionTable}}'] = transactionTableData;
  
  // 15. Autocomplete configuration
  const autocompleteConfig = safeGet(data, 'autocomplete_configuration', null);
  if (autocompleteConfig) {
    textReplacements['{{autocomplete_configuration}}'] = `Enabled: ${formatValue(autocompleteConfig.enabled)}, Sampling Rate: ${formatValue(autocompleteConfig.sampling_rate)}`;
  } else {
    textReplacements['{{autocomplete_configuration}}'] = 'Not configured';
  }
  
  // 16. Limits
  textReplacements['{{max_holdout_campaigns}}'] = formatValue(safeGet(data, 'limits.max_holdout_campaigns'));
  
  // 17. Syndication Settings - EXPANDED
  textReplacements['{{syndication_enabled}}'] = formatValue(safeGet(data, 'syndication_settings.syndication_enabled'));
  textReplacements['{{total_syndications}}'] = formatValue(safeGet(data, 'syndication_settings.total_syndications'));
  textReplacements['{{active_syndications}}'] = formatValue(safeGet(data, 'syndication_settings.active_syndications'));
  
  // Syndication Types Table
  const syndicationTypes = safeGet(data, 'syndication_settings.syndication_types', {});
  const syndicationTypesData = [['Type', 'Count']];
  for (const [type, count] of Object.entries(syndicationTypes)) {
    syndicationTypesData.push([type, formatValue(count)]);
  }
  if (Object.keys(syndicationTypes).length === 0) {
    syndicationTypesData.push(['No syndications', '0']);
  }
  tableReplacements['{{syndication_types_table}}'] = syndicationTypesData;
  
  // 18. Summary Statistics - NEW SECTION
  const summaryStats = safeGet(data, 'summary_statistics', {});
  const summaryData = [['Metric', 'Count']];
  summaryData.push(['Total Campaigns', formatValue(summaryStats.total_campaigns)]);
  summaryData.push(['Total Segments', formatValue(summaryStats.total_segments)]);
  summaryData.push(['Total Templates', formatValue(summaryStats.total_templates)]);
  summaryData.push(['Total Triggers', formatValue(summaryStats.total_triggers)]);
  summaryData.push(['Total Conditions', formatValue(summaryStats.total_conditions)]);
  summaryData.push(['Total Users', formatValue(summaryStats.total_users)]);
  summaryData.push(['Total Catalogs', formatValue(summaryStats.total_catalogs)]);
  summaryData.push(['Total Custom Events', formatValue(summaryStats.total_custom_events)]);
  tableReplacements['{{summary_statistics_table}}'] = summaryData;
  
  return { textReplacements, tableReplacements };
}

// Test function to check permissions and template access
function testPermissions() {
  try {
    // Test DriveApp access
    const files = DriveApp.getFiles();
    console.log('DriveApp access: OK');
    
    // Test DocumentApp access
    const doc = DocumentApp.create('Test Document');
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    console.log('DocumentApp access: OK');
    
    // Test template access
    const templateId = '1Ij4vM7EBqTxJ-TOiHNiWlRUiC-Hf2S8NU4Gull_jWH0';
    try {
      const template = DriveApp.getFileById(templateId);
      console.log('Template access: OK');
      console.log('Template name: ' + template.getName());
    } catch (e) {
      console.error('Template access FAILED: ' + e.toString());
      return 'Cannot access template. Please check the template ID.';
    }
    
    return 'All permissions OK - Template found!';
  } catch (e) {
    console.error('Permission test failed:', e);
    return 'Permission test failed: ' + e.toString();
  }
}

// Helper function to create a sample template
function createSampleTemplate() {
  try {
    // Create a new document to use as template
    const doc = DocumentApp.create('Account Audit Template - ' + new Date().toLocaleDateString());
    const body = doc.getBody();
    
    // Add sample content with placeholders
    body.appendParagraph('Account Audit Report').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    body.appendParagraph('\n');
    body.appendParagraph('1. Account Overview').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Site Name: {{site}}');
    body.appendParagraph('Account ID: {{account_id}}');
    body.appendParagraph('Account Mode: {{account_mode}}');
    body.appendParagraph('Billing Account ID: {{billing_account_id}}');
    body.appendParagraph('Created At: {{created_at}}');
    body.appendParagraph('Timezone: {{timezone}}');
    body.appendParagraph('Currency: {{currency_type}}');
    
    body.appendParagraph('\n');
    body.appendParagraph('2. User Management').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Total Users: {{total_users}}');
    body.appendParagraph('Active Users: {{active_users_count}}');
    body.appendParagraph('Inactive Users: {{inactive_users_count}}');
    body.appendParagraph('\nUsers by Role:');
    body.appendParagraph('{{users_by_role_table}}');
    body.appendParagraph('\nUser Details:');
    body.appendParagraph('{{user_details_table}}');
    
    body.appendParagraph('\n');
    body.appendParagraph('3. Campaign Configuration').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('{{campaignTable}}');
    body.appendParagraph('\nCampaigns by Execution Term:');
    body.appendParagraph('{{campaigns_by_exec_term_table}}');
    body.appendParagraph('\nCampaigns by Status:');
    body.appendParagraph('{{campaigns_by_status_table}}');
    
    body.appendParagraph('\n');
    body.appendParagraph('4. Integration Settings').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Total Integrations: {{total_integrations}}');
    body.appendParagraph('Facebook Status: {{facebook_status}}');
    body.appendParagraph('Google Status: {{google_status}}');
    body.appendParagraph('Webhook Configured: {{webhook_configured}}');
    body.appendParagraph('\nAll Integrations:');
    body.appendParagraph('{{integrations_list_table}}');
    
    body.appendParagraph('\n');
    body.appendParagraph('5. Summary Statistics').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('{{summary_statistics_table}}');
    
    // Save and close
    doc.saveAndClose();
    
    // Get the ID
    const fileId = doc.getId();
    const file = DriveApp.getFileById(fileId);
    
    console.log('Template created successfully!');
    console.log('Template Name: ' + file.getName());
    console.log('Template ID: ' + fileId);
    console.log('Template URL: ' + doc.getUrl());
    
    return 'Template created! ID: ' + fileId + '\nURL: ' + doc.getUrl();
  } catch (e) {
    console.error('Error creating template:', e);
    return 'Error creating template: ' + e.toString();
  }
}