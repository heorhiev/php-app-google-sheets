<?php

namespace app\googleSheet;

use app\toolkit\services\AliasService;
use app\googleSheet\config\GoogleSheetDto;


class GoogleSheet
{
    private $_sheetService;

    private $_options;


    public function __construct(GoogleSheetDto $options)
    {
        $this->_options = $options;

        $client = new \Google_Client();
        $client->setApplicationName($this->_options->appName);
        $client->setScopes([\Google_Service_Sheets::SPREADSHEETS]);
        $client->setAccessType('offline');
        $client->setAuthConfig(AliasService::getAlias($this->_options->apiKeyPath));

        $this->_sheetService = new \Google_Service_Sheets($client);
    }


    public function save(array $rows): bool
    {
        $currentIds = array_flip(array_column($this->getRows(), 0));
        $newIds = array_flip(array_column($rows, 0));

        foreach ($newIds as $newId => $key) {
            if (isset($currentIds[$newId])) {
                $rowNumber = $currentIds[$newId] + 1;
                $this->update($rowNumber, $rows[$key]);

                unset($rows[$key]);
            }
        }

        if ($rows) {
            $this->add($rows);
        }

        return true;
    }


    public function add(array $rows): void
    {
        $valueRange = new \Google_Service_Sheets_ValueRange();
        $valueRange->setValues($rows);

        $this->_sheetService->spreadsheets_values->append(
            $this->_options->sheetId,
            $this->_options->listName,
            $valueRange,
            ['valueInputOption' => 'USER_ENTERED']
        );
    }


    public function update($number, array $row): void
    {
        $rows = [$row];
        $valueRange = new \Google_Service_Sheets_ValueRange();
        $valueRange->setValues($rows);
        $range = $this->_options->listName . '!A' . $number;
        $options = ['valueInputOption' => 'USER_ENTERED'];
        $this->_sheetService->spreadsheets_values->update($this->_options->sheetId, $range, $valueRange, $options);
    }


    public function getRows(): array
    {
        $response = $this->_sheetService->spreadsheets_values->get(
            $this->_options->sheetId,
            $this->_options->listName
        );

        return $response->getValues();
    }
}