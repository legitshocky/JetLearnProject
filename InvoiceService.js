function safeToFixed(value, decimals) {
  if (value === null || value === undefined || isNaN(value)) {
    return '0.00';
  }
  if (decimals === undefined) decimals = 2;
  return Number(value).toFixed(decimals);
}

function getLiveCurrencyRates() {
  // This is our reliable fallback in case the API fails
  const fallbackRates = {
    'EUR': 1.0, 'USD': 1.1700, 'GBP': 0.8665, 'INR': 103.16, 'CHF': 0.9348,
    'AED': 4.273, 'CAD': 1.6224, 'AUD': 1.7682, 'JPY': 172.50, 'SGD': 1.50,
    'HKD': 9.1187, 'ZAR': 20.57, 'CNY': 8.3387, 'NZD': 1.9704, 'SEK': 10.951,
    'NOK': 11.6195, 'DKK': 7.45, 'MXN': 21.8069, 'BRL': 6.3207
  };

  try {
    const API_KEY = PropertiesService.getScriptProperties().getProperty('EXCHANGERATE_API_KEY');
    if (!API_KEY) {
      Logger.log('ExchangeRate API Key not found in Script Properties. Using fallback rates.');
      return fallbackRates;
    }

    const apiUrl = `https://v6.exchangerate-api.com/v6/${API_KEY}/latest/EUR`;
    const response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.result === 'success' && data.conversion_rates) {
        Logger.log('Successfully fetched live currency rates.');
        return data.conversion_rates;
      }
    }
    
    Logger.log('API call for currency rates failed. Response code: ' + response.getResponseCode() + '. Using fallback rates.');
    return fallbackRates;

  } catch (error) {
    Logger.log('Error fetching live currency rates: ' + error.message + '. Using fallback rates.');
    return fallbackRates;
  }
}
function getInvoiceProductsData() {
  Logger.log('getInvoiceProductsData called');
  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.INVOICE_PRODUCTS);
    if (sheetData.length < 2) { 
      Logger.log('Invoice Products sheet is empty or only has headers.');
      return [];
    }
    const headers = sheetData[0];
    const products = sheetData.slice(1).map(row => {
      const product = {};
      headers.forEach((header, i) => {
        product[header.trim()] = row[i];
      });
      return product;
    }).filter(p => p['Plan Name'] !== '');
    Logger.log(`Found ${products.length} invoice products.`);
    return products;
  } catch (error) {
    Logger.log('Error getting invoice products data: ' + error.message);
    return [];
  }
}
function getLiveExchangeRates() {
  Logger.log('Fetching live exchange rates from API.');
  try {
    const API_KEY = PropertiesService.getScriptProperties().getProperty('EXCHANGERATE_API_KEY');
    if (!API_KEY) {
      Logger.log('ExchangeRate API Key not found in Script Properties.');
      return {};
    }
    const apiUrl = `https://v6.exchangerate-api.com/v6/${API_KEY}/latest/EUR`;
    const response = UrlFetchApp.fetch(apiUrl, {
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      Logger.log(`Exchange rate API returned non-200 status: ${responseCode}. Response: ${response.getContentText()}`);
      return {}; 
    }

    const jsonResponse = JSON.parse(response.getContentText());

    if (jsonResponse && jsonResponse.result === 'success' && jsonResponse.conversion_rates) {
      const rates = jsonResponse.conversion_rates;
      
      delete rates.EUR;
      delete rates.GBP;
      delete rates.USD;

      Logger.log(`Successfully fetched and filtered ${Object.keys(rates).length} live exchange rates.`);
      return rates;
    } else {
      Logger.log(`Exchange rate API response was invalid or unsuccessful. Response: ${JSON.stringify(jsonResponse)}`);
      return {}; 
    }
  } catch (error) {
    Logger.log(`Error fetching or parsing live exchange rates: ${error.message}. Stack: ${error.stack}`);
    return {}; 
  }
}
function getConversionRate(toCurrency) {
  toCurrency = toCurrency.toUpperCase();
  if (toCurrency === 'EUR') {
    return 1.0;
  }

  if (!_sheetDataCache['liveRates']) {
    Logger.log('Fetching live currency rates for this execution...');
    _sheetDataCache['liveRates'] = getLiveCurrencyRates(); 
  }
  const liveRates = _sheetDataCache['liveRates'];

  if (liveRates && liveRates[toCurrency]) {
    Logger.log(`Using LIVE rate for EUR to ${toCurrency}: ${liveRates[toCurrency]}`);
    return liveRates[toCurrency];
  }
  
  const fallbackRates = {
    'EUR': 1.0, 'USD': 1.1700, 'GBP': 0.8665, 'INR': 103.16, 'CHF': 0.9348,
    'AED': 4.273, 'CAD': 1.6224, 'AUD': 1.7682, 'JPY': 172.50, 'SGD': 1.50,
    'HKD': 9.1187, 'ZAR': 20.57, 'CNY': 8.3387, 'NZD': 1.9704, 'SEK': 10.951,
    'NOK': 11.6195, 'DKK': 7.45, 'MXN': 21.8069, 'BRL': 6.3207
  };
  
  const fallbackRate = fallbackRates[toCurrency];
  if (fallbackRate !== undefined) {
    Logger.log(`Using FALLBACK hardcoded rate for EUR to ${toCurrency}: ${fallbackRate}`);
    return fallbackRate;
  }

  Logger.log(`No live or hardcoded rate found for EUR to ${toCurrency}. Returning 1 as a safe default.`);
  return 1; 
}

function getCurrencySymbol(currencyCode) {
    switch (currencyCode) {
        case 'GBP': return '£';
        case 'EUR': return '€';
        case 'USD': return '$';
        case 'INR': return '₹';
        case 'JPY': return '¥';
        case 'AUD': return 'A$';
        case 'CAD': return 'C$';
        case 'CHF': return 'CHF';
        case 'CNY': return '¥';
        case 'SEK': return 'kr';
        case 'NZD': return 'NZ$';
        case 'AED': return 'د.إ';
        case 'HKD': return 'HK$';
        case 'ZAR': return 'R';
        case 'SGD': return 'S$';
        case 'NOK': return 'kr';
        case 'DKK': return 'kr';
        case 'MXN': return 'Mex$';
        case 'BRL': return 'R$';
        default: return currencyCode || ''; 
    }
}

function calculateInvoicePricing(formData, previewOnly = false) {
    let effectiveBasePrice = 0;
    let discount = parseFloat(formData.discount || '0');
    let customCurrencyExtraDiscountPercentage = parseFloat(formData.customCurrencyExtraDiscountPercentage || '0');
    let finalCurrencySymbol = getCurrencySymbol(formData.currency); 

    const userSelectedTenureMonths = parseInt(formData.subscriptionTenure || '0');
    const freeClasses = parseInt(formData.freeClasses || '0');
    const sessionsPerWeekNum = parseInt((formData.sessionsPerWeek && String(formData.sessionsPerWeek).split(' ')[0]) || '0');
    
    // We treat this as "Amount Paid So Far"
    const customPaidAmount = parseFloat(formData.customPaidAmount || '0');
    
    const invoiceProducts = getInvoiceProductsData();
    const selectedPlan = invoiceProducts.find(p => p['Plan Name'] === formData.planName);

    if (!selectedPlan) {
        throw new Error(`Invoice plan '${formData.planName}' not found.`);
    }
    
    let fixedClassesPerPlan = parseInt(selectedPlan['Fixed Classes']);
    if (isNaN(fixedClassesPerPlan)) fixedClassesPerPlan = 0;

    let targetCurrencyCode = formData.currency; 
    let finalConversionRate = 1.0; 

    if (targetCurrencyCode === 'CUSTOM') {
        targetCurrencyCode = (formData.customCurrencyCode || 'EUR').toUpperCase(); 
        finalConversionRate = (formData.customCurrencyRate && parseFloat(formData.customCurrencyRate) > 0) 
            ? parseFloat(formData.customCurrencyRate) 
            : getConversionRate(targetCurrencyCode);
        finalCurrencySymbol = getCurrencySymbol(targetCurrencyCode);
    } else if (!['EUR', 'USD', 'GBP'].includes(targetCurrencyCode)) {
        finalConversionRate = getConversionRate(targetCurrencyCode);
    }
    
    let planBasePriceForDefaultTenure = 0;
    if (formData.currency === 'USD') {
        planBasePriceForDefaultTenure = parseFloat(String(selectedPlan['Base Price USD'] || '0').replace(/[^0-9.-]/g, '')) || 0;
    } else if (formData.currency === 'GBP') {
        planBasePriceForDefaultTenure = parseFloat(String(selectedPlan['Base Price GBP'] || '0').replace(/[^0-9.-]/g, '')) || 0;
    } else {
        planBasePriceForDefaultTenure = parseFloat(String(selectedPlan['Base Price EUR'] || '0').replace(/[^0-9.-]/g, '')) || 0;
    }

    const selectedPlanDefaultMonthsTenure = parseInt(selectedPlan['Months Tenure'] || '1');
    const monthlyRate = selectedPlanDefaultMonthsTenure > 0 ? (planBasePriceForDefaultTenure / selectedPlanDefaultMonthsTenure) : 0;
    
    effectiveBasePrice = monthlyRate * userSelectedTenureMonths * finalConversionRate;
    if (isNaN(effectiveBasePrice)) effectiveBasePrice = 0;

    if (formData.currency === 'CUSTOM' && effectiveBasePrice > 0 && customCurrencyExtraDiscountPercentage > 0) {
        effectiveBasePrice *= (1 - customCurrencyExtraDiscountPercentage / 100);
    }

    // --- 1. Total Calculation Fix ---
    // Total is ALWAYS Base - Discount
    let finalTotal = effectiveBasePrice - discount;
    if (finalTotal < 0) finalTotal = 0;

    const numInstallments = parseInt(formData.numberOfInstallments || '1');
    const isInstallment = formData.paymentType === 'Installment' && numInstallments > 0;

    // --- 2. Amount Paid & Balance Logic ---
    let amountPaid = 0;
    
    if (customPaidAmount > 0) {
        // Explicit partial payment entered
        amountPaid = customPaidAmount;
    } else {
        // Auto-calculate
        if (isInstallment && numInstallments > 0) {
            amountPaid = finalTotal / numInstallments;
        } else {
            amountPaid = finalTotal;
        }
    }
    
    const balanceDue = Math.max(0, finalTotal - amountPaid);
    // -------------------------------------

    let totalClasses = (fixedClassesPerPlan > 0) ? fixedClassesPerPlan + freeClasses : (userSelectedTenureMonths * 4 * sessionsPerWeekNum) + freeClasses;
    totalClasses = Math.max(0, totalClasses);
    
    const unitPrice = totalClasses > 0 ? (effectiveBasePrice / totalClasses) : 0;

    let weeksRequired = (sessionsPerWeekNum > 0) ? Math.ceil(totalClasses / sessionsPerWeekNum) : 0;
    
    let finalEndDate;
    if (formData.endDate && formData.endDate.trim() !== '') {
        finalEndDate = new Date(formData.endDate);
    } else {
        const startDate = new Date(formData.startDate);
        finalEndDate = new Date(startDate);
        if (weeksRequired > 0) {
            finalEndDate.setDate(startDate.getDate() + (weeksRequired * 7) - 1);
        }
    }

    // --- 3. Installment Array Logic ---
    const installments = [];
    
    if (isInstallment && formData.installmentType && formData.dueDayToPay) {
        const billingAnchorDate = formData.firstPaymentDate ? new Date(formData.firstPaymentDate) : new Date(formData.startDate);
        const dueDay = parseInt(formData.dueDayToPay);
        
        let monthIncrement = 1;
        switch (formData.installmentType) {
            case 'Alternate': monthIncrement = 2; break;
            case 'Quarterly': monthIncrement = 3; break;
            case '4 Months': monthIncrement = 4; break;
            case '5 Months': monthIncrement = 5; break;
            case '6 Months': monthIncrement = 6; break;
            case '7 Months': monthIncrement = 7; break;
            case '8 Months': monthIncrement = 8; break;
            case '9 Months': monthIncrement = 9; break;
            case '10 Months': monthIncrement = 10; break;
            case '11 Months': monthIncrement = 11; break;
            case '12 Months': monthIncrement = 12; break;
        }

        if (customPaidAmount > 0) {
            // CASE A: Custom Partial Payment Logic
            // 1. First payment is what they entered (Is Paid: YES)
            installments.push({ 
                number: 1, 
                amount: customPaidAmount, 
                isPaid: true, 
                dueDate: billingAnchorDate, 
                dueDateFormatted: formatDateDDMMYYYY(billingAnchorDate) 
            });

            // 2. Remaining installments split the Balance Due
            const remainingCount = numInstallments - 1; 
            
            if (remainingCount > 0 && balanceDue > 0) {
                const nextAmount = balanceDue / remainingCount;
                let lastDueDate = new Date(billingAnchorDate);

                for (let i = 0; i < remainingCount; i++) {
                    let nextDueDate = new Date(lastDueDate);
                    nextDueDate.setMonth(nextDueDate.getMonth() + monthIncrement);
                    nextDueDate.setDate(dueDay);
                    
                    installments.push({ 
                        number: i + 2, 
                        amount: nextAmount, 
                        isPaid: false, 
                        dueDate: nextDueDate, 
                        dueDateFormatted: formatDateDDMMYYYY(nextDueDate) 
                    });
                    lastDueDate = nextDueDate;
                }
            }
        } else {
            // CASE B: Standard Logic (Even Split)
            const installmentAmount = finalTotal / numInstallments;
            
            installments.push({ 
                number: 1, 
                amount: installmentAmount, 
                isPaid: true, 
                dueDate: billingAnchorDate, 
                dueDateFormatted: formatDateDDMMYYYY(billingAnchorDate) 
            });

            let lastDueDate = new Date(billingAnchorDate);
            for (let i = 1; i < numInstallments; i++) {
                let nextDueDate = new Date(lastDueDate);
                nextDueDate.setMonth(nextDueDate.getMonth() + monthIncrement);
                nextDueDate.setDate(dueDay);
                
                installments.push({ 
                    number: i + 1, 
                    amount: installmentAmount, 
                    isPaid: false, 
                    dueDate: nextDueDate, 
                    dueDateFormatted: formatDateDDMMYYYY(nextDueDate) 
                });
                lastDueDate = nextDueDate;
            }
        }
    }

    return {
        unitPrice: unitPrice,
        effectiveBasePrice: effectiveBasePrice,
        discount: discount,
        finalTotal: finalTotal, // Total Deal Value
        amountPaid: amountPaid, // Collected
        balanceDue: balanceDue, // Remaining
        currencySymbol: finalCurrencySymbol,
        subscriptionTenureMonths: userSelectedTenureMonths,
        startDateFormatted: formatDate(new Date(formData.startDate)),
        endDateFormatted: formatDate(finalEndDate),
        paymentType: formData.paymentType,
        numberOfInstallments: numInstallments,
        displayTotalSessions: totalClasses,
        planDescription: formData.planName,
        installments: installments,
        upfrontDueDate: formData.upfrontDueDate ? formatDate(new Date(formData.upfrontDueDate)) : null
    };
}

function validateInvoiceData(formData) { 
  const errors = [];

  if (!formData.learnerName || formData.learnerName.trim() === '') errors.push('Learner Name is required');
  if (formData.learnerEmail && !isValidEmail(formData.learnerEmail)) errors.push('Learner Email, if provided, must be valid.');

  if (!formData.parentName || formData.parentName.trim() === '') errors.push('Parent Name is required');
  if (!formData.parentEmail || !isValidEmail(formData.parentEmail)) errors.push('Valid Parent Email is required');
  if (!formData.parentContact || formData.parentContact.trim() === '') errors.push('Parent Contact is required');
  if (!formData.planName || formData.planName.trim() === '') errors.push('Plan Name is required');
  if (!formData.currency || formData.currency.trim() === '') errors.push('Currency is required');
  if (!formData.sessionsPerWeek || formData.sessionsPerWeek.trim() === '') errors.push('Sessions per Week is required');
  if (parseInt(formData.subscriptionTenure || '0') < 0) errors.push('Subscription Tenure must be a non-negative number.');
  if (!formData.startDate || formData.startDate.trim() === '') errors.push('Start Date is required');
  if (!formData.endDate || formData.endDate.trim() === '') errors.push('End Date is required'); 
  if (new Date(formData.endDate) < new Date(formData.startDate)) errors.push('End Date cannot be before Start Date.');
  if (!formData.paymentType || formData.paymentType.trim() === '') errors.push('Payment Type is required');

  let discount = parseFloat(formData.discount || '0'); 
  const customPaidAmount = parseFloat(formData.customPaidAmount || '0'); 
  const freeClasses = parseInt(formData.freeClasses || '0'); 

  if (discount < 0) errors.push('Discount cannot be negative.');
  if (customPaidAmount < 0) errors.push('Partial Payment Received cannot be negative.');
  if (freeClasses < 0) errors.push('Free Classes cannot be negative.'); 

  const invoiceProducts = getInvoiceProductsData();
  const selectedPlan = invoiceProducts.find(p => p['Plan Name'] === formData.planName);

  if (selectedPlan && formData.currency && (parseInt(formData.subscriptionTenure || '0') >= 0)) {
      let targetCurrencyCode = formData.currency;
      let finalConversionRateFromEUR = 1;
      let customCurrencyExtraDiscountPercentage = parseFloat(formData.customCurrencyExtraDiscountPercentage || '0');

      if (targetCurrencyCode === 'CUSTOM') {
          if (!formData.customCurrencyCode || formData.customCurrencyCode.trim() === '') errors.push('Custom Currency Code is required for custom currency.');
          targetCurrencyCode = (formData.customCurrencyCode || 'EUR').toUpperCase(); 
          finalConversionRateFromEUR = (formData.customCurrencyRate && parseFloat(formData.customCurrencyRate) > 0)
            ? parseFloat(formData.customCurrencyRate)
            : getConversionRate(targetCurrencyCode);

          if (finalConversionRateFromEUR <= 0) {
              errors.push(`Invalid conversion rate from EUR to ${targetCurrencyCode}.`);
          }
          if (customCurrencyExtraDiscountPercentage < 0 || customCurrencyExtraDiscountPercentage > 100) {
              errors.push('Custom currency discount percentage must be between 0 and 100.');
          }
      } else {
          finalConversionRateFromEUR = getConversionRate(targetCurrencyCode);
      }

      const planBasePriceForDefaultTenure = parseFloat(String(selectedPlan['Base Price EUR'] || '0').replace(/[^0-9.-]/g, '')) || 0;
      const selectedPlanDefaultMonthsTenure = parseInt(selectedPlan['Months Tenure'] || '1'); 
      const userSelectedTenureMonths = parseInt(formData.subscriptionTenure || formData.subscriptionTenureMonths || '0');

      const monthlyRateEUR = selectedPlanDefaultMonthsTenure > 0 ? (planBasePriceForDefaultTenure / selectedPlanDefaultMonthsTenure) : 0;
      let effectiveBasePrice = monthlyRateEUR * userSelectedTenureMonths;
      if (isNaN(effectiveBasePrice)) effectiveBasePrice = 0;

      if (formData.currency === 'CUSTOM' && effectiveBasePrice > 0 && customCurrencyExtraDiscountPercentage > 0) {
          effectiveBasePrice *= (1 - customCurrencyExtraDiscountPercentage / 100);
      }
      effectiveBasePrice *= finalConversionRateFromEUR;

      // if (discount > effectiveBasePrice) {
      //     errors.push('Discount cannot exceed the calculated base price. It will be capped automatically.');
      // }
  }

  if (formData.paymentType === 'Installment') {
      const numInstallments = parseInt(formData.numberOfInstallments || '0');
      if (isNaN(numInstallments) || numInstallments <= 0) {
          errors.push('Number of Installments must be a positive number for installment plans.');
      }
      if (!formData.installmentType || formData.installmentType.trim() === '') {
          errors.push('Installment Type is required for installment plans.');
      }
      const dueDay = parseInt(formData.dueDayToPay || '0');
      if (isNaN(dueDay) || dueDay < 1 || dueDay > 31) {
          errors.push('Due Day to Pay must be a number between 1 and 31 for installment plans.');
      }
  }

  return errors;
}
function generateInvoicePDFAndEmail(formData) { 
  Logger.log(`generateInvoicePDFAndEmail called for Learner: ${formData.learnerName}`);
  let trackingId = null; // To capture the ID for logging, even on failure
  
  try {
    const validationErrors = validateInvoiceData(formData); 
    if (validationErrors.length > 0) {
      throw new Error('Validation failed: ' + validationErrors.join(', '));
    }

    const pricingDetails = calculateInvoicePricing(formData); 
    const invoiceHtml = getInvoiceHTML(formData, pricingDetails);

    const pdfName = `Invoice-${formData.learnerName.replace(/\s/g, '_')}-${formData.jlid || 'N_A'}-${new Date().toISOString().split('T')[0]}.pdf`;
    const blob = Utilities.newBlob(invoiceHtml, 'text/html')
                             .getAs(MimeType.PDF)
                             .setName(pdfName);

    // Save to Google Drive
    if (CONFIG.DRIVE_FOLDER_ID) {
        try {
            DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID).createFile(blob);
            Logger.log(`Invoice PDF saved to Drive: ${pdfName}`);
        } catch (e) {
            Logger.log(`Warning: Could not save invoice PDF to Drive: ${e.message}`);
            // Non-fatal error, we can still try to email it.
        }
    }

    // Prepare the simple wrapper email body for tracking
    const emailHtmlBody = `
     <!DOCTYPE html>
      <html>
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>JetLearn Invoice</title>
      </head>
      <body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f5f5f5;">
          <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f5f5f5; padding: 20px;">
              <tr>
                  <td align="center">
                      <table width="600" cellpadding="0" cellspacing="0" style="background-color: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
                          
                          <!-- Header -->
                          <tr>
                              <td style="background: linear-gradient(135deg, #FFD700 0%, #FFA500 100%); padding: 30px 40px; text-align: center;">
                                  <h1 style="margin: 0; color: #000; font-size: 32px; font-weight: 700; letter-spacing: 0.5px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
                                      JetLearn
                                  </h1>
                                  <p style="margin: 8px 0 0; color: #333; font-size: 14px; font-weight: 500;">
                                      World's Top Online AI Academy
                                  </p>
                              </td>
                          </tr>
                          
                          <!-- Content -->
                          <tr>
                              <td style="padding: 40px;">
                                  <p style="color: #666; line-height: 1.6; margin: 0 0 20px; font-size: 16px;">
                                      Dear ${formData.parentName},
                                  </p>
                                  
                                  <p style="color: #666; line-height: 1.6; margin: 0 0 20px; font-size: 16px;">
                                      Please find attached the invoice for <strong style="color: #000;">${formData.learnerName}</strong>'s reference.
                                  </p>
                                  
                                  <p style="color: #666; line-height: 1.6; margin: 0 0 20px; font-size: 16px;">
                                      We appreciate you giving your child the best opportunity to learn and grow. We are honored to be part of this journey!
                                  </p>
                                  
                                  <p style="color: #666; line-height: 1.6; margin: 25px 0 0; font-size: 16px;">
                                      Best regards,<br>
                                      <strong style="color: #000;">The JetLearn Team</strong>
                                  </p>
                              </td>
                          </tr>
                          
                          <!-- Footer -->
                          <tr>
                              <td style="background-color: #000; padding: 25px 40px; text-align: center;">
                                  <p style="margin: 0; color: #FFD700; font-size: 12px;">
                                      © 2025 JetLearn. Empowering kids to lead in the age of AI.
                                  </p>
                              </td>
                          </tr>
                          
                      </table>
                  </td>
              </tr>
          </table>
      </body>
      </html>
    `;

    // Use the central tracked email service
    const emailResult = sendTrackedEmail({
      to: formData.parentEmail,
      subject: `JetLearn Invoice for ${formData.learnerName}`,
      htmlBody: emailHtmlBody,
      jlid: formData.jlid,
      attachments: [blob]
    });
    trackingId = emailResult.trackingId;

    // Log the successful action to the main audit log
    logAction('Invoice Sent', formData.jlid || '', formData.learnerName, '', '', formData.planName, 'Success', `Invoice PDF sent to ${formData.parentEmail}. TID: ${trackingId}`);
    
    return { success: true, message: 'Invoice generated and emailed successfully!' };

  } catch (error) {
    Logger.log('Error in generateInvoicePDFAndEmail: ' + error.message);
    // Log the failed action to the main audit log
    logAction('Invoice Failed', formData.jlid || '', formData.learnerName, '', '', formData.planName, 'Failed', `Error: ${error.message}. TID Attempt: ${trackingId}`);
    return { success: false, message: 'Failed to send invoice email: ' + error.message };
  } 
}
function generateInvoicePDFForDownload(formData) { 
  Logger.log(`generateInvoicePDFForDownload called for Learner: ${formData.learnerName}`);

  const validationErrors = validateInvoiceData(formData); 
  if (validationErrors.length > 0) {
    return { success: false, message: 'Validation failed: ' + validationErrors.join(', ')} ;
  }

  if (!formData.learnerEmail || formData.learnerEmail.trim() === '') {
      formData.learnerEmail = formData.parentEmail;
  }

  const pricingDetails = calculateInvoicePricing(formData, false); 

  const htmlBody = getInvoiceHTML(formData, pricingDetails);

  const pdfName = `Invoice-${formData.learnerName.replace(/\s/g, '_')}-${formData.jlid || 'N_A'}-${new Date().toISOString().split('T')[0]}_${Date.now()}.pdf`;
  const blob = Utilities.newBlob(htmlBody, 'text/html')
                           .getAs(MimeType.PDF)
                           .setName(pdfName);

  if (!CONFIG.DRIVE_FOLDER_ID || CONFIG.DRIVE_FOLDER_ID === '1_exampleFolderID1234567890abcdef') {
      Logger.log("DRIVE_FOLDER_ID is not configured. Invoice PDF will not be saved to Drive.");
  } else {
      try {
          const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
          folder.createFile(blob);
          Logger.log(`Invoice PDF saved to Drive for download: ${pdfName}`);
      } catch (e) {
          Logger.log(`Error saving invoice PDF to Drive for download: ${e.message}`);
      }
  }

  try {
      const base64Data = Utilities.base64Encode(blob.getBytes());
      return { success: true, message: 'PDF generated successfully for download.', filename: pdfName, data: base64Data };
  } catch (error) {
      Logger.log('Error encoding PDF to base64: ' + error.message);
      return { success: false, message: 'Failed to encode PDF for download: ' + error.message };
  }
}

function sendRenewalCommunication(formData) {
  try {
    // 1. Generate PDF
    const template = HtmlService.createTemplateFromFile('RenewalInvoiceTemplate');
    template.data = formData;
    const html = template.evaluate().getContent();
    const pdfBlob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF)
                             .setName(`Renewal_Invoice_${formData.learnerName}.pdf`);

    // 2. Prepare Email HTML
    const emailTemplate = HtmlService.createTemplateFromFile('RenewalEmailTemplate');
    emailTemplate.data = formData;
    const emailBody = emailTemplate.evaluate().getContent();

    // 3. Send via central tracker
    sendTrackedEmail({
        to: formData.parentEmail,
        subject: `Renewal Confirmation - ${formData.learnerName}`,
        htmlBody: emailBody,
        jlid: formData.jlid,
        attachments: [pdfBlob]
    });

    logAction('Renewal Sent', formData.jlid, formData.learnerName, '', '', formData.planName, 'Success', `Renewal Amount: ${formData.netPrice}`);

    return { success: true };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getInvoiceHTML(formData, pricingDetails) {
  const jlid = (formData.jlid || '').trim().toUpperCase();
  let planDescription = 'Comprehensive Learning Program'; 

  if (jlid.endsWith('C')) {
    planDescription = 'Comprehensive AI Coding Program';
  } else if (jlid.endsWith('M')) {
    planDescription = 'Comprehensive Math Program';
  }

  formData.planDescription = planDescription;

  const template = HtmlService.createTemplateFromFile('InvoiceTemplate');
  template.data = formData;
  template.pricing = pricingDetails; 
  return template.evaluate().getContent();
}


// Helpers for Preview Generation
function getRenewalEmailHTML(formData) {
  const template = HtmlService.createTemplateFromFile('RenewalEmailTemplate');
  template.data = formData;
  return template.evaluate().getContent();
}

function getRenewalInvoiceHTML(formData) {
  const template = HtmlService.createTemplateFromFile('RenewalInvoiceTemplate');
  template.data = formData;
  return template.evaluate().getContent();
}
