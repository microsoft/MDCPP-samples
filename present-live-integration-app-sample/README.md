# Present-live Sample Integration App

The present-live experience is currently available for dev preview in the [MDCPP npm package](https://aka.ms/MDCPP-npm-package). This sample app demonstrates how to use the MDCPP npm package to integrate this experience.

The present-live experience allows meeting participants to interact and collaborate using PowerPoint Live directly in your collaboration app. For more about the present-live experience and how to implement it in your app, see [Integrate present-live experiences in the Microsoft 365 Document Collaboration Partner Program](https://learn.microsoft.com/microsoft-365/document-collaboration-partner-program/scenarios/present).

> **IMPORTANT**:
>
> - You must be a member of the Microsoft 365 Document Collaboration Partner Program in order to access MDCPP services.
> - For the dev preview, don't use any Microsoft 365 accounts that contain real user data. Data for this experience is being hosted on development endpoints.

## Set up

Follow these instructions to configure a sample app in your test tenant.

1. Sign in to the Entra admin center with your test tenant admin account: https://entra.microsoft.com/
1. Expand the **Applications** tab and open the **App registrations** page.

    ![App registrations link on Microsoft Entra admin center's panel.](images/admin-center-registrations-tab.png)

1. Select **New registration** and give your sample app a name.
1. Select **Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant)** under **Supported account types**.
1. Add an SPA redirect URI pointing to http://localhost/desktop.html as shown in the following screenshot.

   ![Register an application.](images/entra-new-app-registration.png)

1. Select **Register**.
1. Choose the **API permissions** tab on your sample app's registration page.
1. Select **Add a permission** > **SharePoint** > **Delegated permissions**.
1. Ensure that **AllSites.Read** and **MyFiles.Read** permissions are selected. Select **Add permissions**.
1. Select **Grant admin consent for \<tenant-name\>** (where \<tenant-name\> is the name of your test tenant) then choose **Yes** in the subsequent popup dialog. Your registration page should look similar to the following screenshot.

   ![Tenant application registered.](images/entra-tenant-app-registered.png)

1. Ensure that your test user account's OneDrive files contain at least one PowerPoint slide deck for testing.

   > **Note**: You can create a sample PowerPoint presentation file at https://m365.cloud.microsoft/create.

## Run the sample app locally

1. Fork this repo then clone your fork locally.
1. In the repo's root directory, create a new file named **.env**.
1. Open the .env file in your preferred code editor.
1. Add the following line to the .env file, replacing "[your Entra client ID]" with your Entra sample app's client ID. You can find the ID by navigating to the **Overview** tab in your app's registration page, listed under **Application (client) ID** in the **Essentials** dropdown menu near the top of the page.

   ```text
   ENTRA_APPID = [your Entra client ID]
   ```

1. In a terminal, run the following command from the repo's root directory to download dependencies.

   ```text
   yarn
   ```

   1. If you don't have yarn, you can download it by running the following command.

      ```text
      npm i -g yarn
      ```

   1. If you don't have Node.js, you can download it from https://nodejs.org/en/download.

1. Run the following command to launch the sample app. Ensure that port 8080 is free.

   ```text
   yarn start
   ```

1. In a private or incognito browser window, navigate to the following URL. To avoid issues with signing in, it's recommended to do this in a private or incognito browser window in case you're already signed in with a different Entra account in that browser.

   ```text
   http://localhost:8080/desktop.html
   ```

1. Sign in with your test Entra account. You may need to accept a **Permissions requested** dialog. After a few seconds, the file picker UI should load.
1. Select your test PowerPoint file from the file picker.
1. Choose **Present** at the bottom right.

A new browser tab should open that contains the attendee view. The initial browser tab will contain the presenter view. You may need to allow browser popups to see the new attendee tab.
