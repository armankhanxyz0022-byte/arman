<?php
// index.php - Professional Business Dashboard (No Sidebar)
// Meesho / Flipkart / Amazon / Myntra / Auto-detect

error_reporting(E_ALL);
ini_set('display_errors', 0);
ini_set('memory_limit', '512M');
set_time_limit(300);

$uploadDir  = __DIR__ . '/uploads/';
$extractDir = __DIR__ . '/uploads/extracted/';
if (!is_dir($uploadDir))  mkdir($uploadDir, 0777, true);
if (!is_dir($extractDir)) mkdir($extractDir, 0777, true);

$summary      = null;
$error        = '';
$debug        = [];
$fileResults  = [];
$orderDetails = [];
$selectedPlatform = $_POST['platform'] ?? 'auto';

// ================================================================
//  NATIVE XLSX READER
// ================================================================
class NativeXlsxReader
{
    private $zip;
    private $sharedStrings = [];
    private $sheetMap      = [];

    public function load($path)
    {
        if (!class_exists('ZipArchive')) throw new Exception('PHP Zip extension not enabled.');
        $this->zip = new ZipArchive();
        if ($this->zip->open($path) !== true) throw new Exception('Cannot open XLSX.');
        $this->sharedStrings = [];
        $ssXml = $this->zip->getFromName('xl/sharedStrings.xml');
        if ($ssXml !== false) $this->parseSharedStrings($ssXml);
        $this->sheetMap = $this->buildSheetMap();
    }

    private function parseSharedStrings($xml)
    {
        $reader = new XMLReader();
        $reader->XML($xml);
        $currentString = ''; $inSi = false;
        while ($reader->read()) {
            if ($reader->nodeType === XMLReader::ELEMENT) {
                if ($reader->localName === 'si') { $inSi = true; $currentString = ''; }
                if ($reader->localName === 't' && $inSi) $currentString .= $reader->readString();
            }
            if ($reader->nodeType === XMLReader::END_ELEMENT) {
                if ($reader->localName === 'si') {
                    $this->sharedStrings[] = $currentString;
                    $inSi = false;
                }
            }
        }
        $reader->close();
    }

    private function buildSheetMap()
    {
        $map = [];
        $wbXml = $this->zip->getFromName('xl/workbook.xml');
        if ($wbXml === false) return $map;
        $relsXml = $this->zip->getFromName('xl/_rels/workbook.xml.rels');
        $rels = [];
        if ($relsXml !== false) {
            if (preg_match_all('/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/i', $relsXml, $m)) {
                foreach ($m[1] as $i => $rid) {
                    $target = ltrim($m[2][$i], '/');
                    if (strpos($target, 'xl/') !== 0) $target = 'xl/' . $target;
                    $rels[$rid] = $target;
                }
            }
        }
        if (preg_match_all('/<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/i', $wbXml, $sm)) {
            foreach ($sm[1] as $i => $name) {
                $rid = $sm[2][$i];
                if (isset($rels[$rid])) $map[$name] = $rels[$rid];
            }
        }
        return $map;
    }

    public function getSheetNames() { return array_keys($this->sheetMap); }

    public function getAllSheets()
    {
        $out = [];
        foreach ($this->sheetMap as $name => $path) $out[$name] = $this->parseSheet($path);
        return $out;
    }

    public function getRows()
    {
        $best = []; $bestCount = 0;
        foreach ($this->sheetMap as $name => $path) {
            $rows = $this->parseSheet($path);
            $nonEmpty = 0;
            foreach ($rows as $r) {
                if (count(array_filter($r, fn($v) => $v !== null && $v !== '')) >= 2) $nonEmpty++;
            }
            if ($nonEmpty > $bestCount) { $bestCount = $nonEmpty; $best = $rows; }
        }
        return $best;
    }

    private function parseSheet($sheetPath)
    {
        $xml = $this->zip->getFromName($sheetPath);
        if ($xml === false) return [];
        $rows = [];
        $reader = new XMLReader();
        $reader->XML($xml);
        $currentRow = []; $colIndex = -1; $inCell = false;
        $cellType = ''; $cellValue = '';
        while ($reader->read()) {
            if ($reader->nodeType === XMLReader::ELEMENT) {
                if ($reader->localName === 'row') $currentRow = [];
                elseif ($reader->localName === 'c') {
                    $inCell = true;
                    $cellType = $reader->getAttribute('t') ?? '';
                    $ref = $reader->getAttribute('r') ?? '';
                    $colIndex = $this->colLetterToIndex($ref);
                    $cellValue = '';
                } elseif ($inCell && $reader->localName === 'v') $cellValue = $reader->readString();
                elseif ($inCell && $reader->localName === 't') $cellValue = $reader->readString();
            } elseif ($reader->nodeType === XMLReader::END_ELEMENT) {
                if ($reader->localName === 'c' && $inCell) {
                    $value = ($cellType === 's') ? ($this->sharedStrings[(int)$cellValue] ?? '') : $cellValue;
                    if ($colIndex >= 0) $currentRow[$colIndex] = $value;
                    $inCell = false; $colIndex = -1;
                } elseif ($reader->localName === 'row') {
                    if (!empty($currentRow)) {
                        $maxCol = max(array_keys($currentRow));
                        $norm = [];
                        for ($c = 0; $c <= $maxCol; $c++) $norm[$c] = $currentRow[$c] ?? '';
                        $rows[] = $norm;
                    }
                }
            }
        }
        $reader->close();
        return $rows;
    }

    private function colLetterToIndex($ref)
    {
        preg_match('/^([A-Z]+)/', $ref, $m);
        if (empty($m[1])) return -1;
        $letters = $m[1]; $num = 0;
        for ($i = 0; $i < strlen($letters); $i++) $num = $num * 26 + (ord($letters[$i]) - ord('A') + 1);
        return $num - 1;
    }

    public function close() { if ($this->zip) $this->zip->close(); }
}

// ========== HELPERS ==========
function cleanNumber($val) {
    if (is_numeric($val)) return (float)$val;
    $val = str_replace([',', '₹', '$', ' ', '"', "\xC2\xA0"], '', (string)$val);
    return is_numeric($val) ? (float)$val : 0;
}

function detectStatus($s) {
    $s = strtolower(trim($s));
    if ($s === '') return 'unknown';
    if (strpos($s, 'rto') !== false || strpos($s, 'return to origin') !== false) return 'rto';
    if (strpos($s, 'return') !== false || strpos($s, 'refund') !== false) return 'return';
    if (strpos($s, 'cancel') !== false) return 'cancelled';
    if (strpos($s, 'delivered') !== false) return 'delivered';
    if (strpos($s, 'shipped') !== false || strpos($s, 'dispatch') !== false || strpos($s, 'intransit') !== false || strpos($s, 'in_transit') !== false) return 'shipped';
    if (strpos($s, 'pending') !== false || strpos($s, 'processing') !== false || strpos($s, 'approved') !== false) return 'pending';
    return 'other';
}

function meeshoCols() {
    return [
        'sub_order_no'=>0,'order_date'=>1,'dispatch_date'=>2,'product'=>3,
        'sku'=>4,'catalog_id'=>5,'order_source'=>6,'status'=>7,
        'product_gst'=>8,'listing_price'=>9,'qty'=>10,'transaction_id'=>11,
        'payment_date'=>12,'final_settlement'=>13,'price_type'=>14,
        'total_sale_amount'=>15,'total_sale_return'=>16,'fixed_fee_1'=>17,
        'warehousing_fee_1'=>18,'return_premium'=>19,'return_premium_return'=>20,
        'commission_pct'=>21,'meesho_commission'=>22,'meesho_gold_fee'=>23,
        'meesho_mall_fee'=>24,'fixed_fee_2'=>25,'warehousing_fee_2'=>26,
        'return_shipping'=>27,'gst_compensation'=>28,'shipping_charge'=>29,
        'other_support'=>30,'waivers'=>31,'net_other_support'=>32,
        'gst_on_other'=>33,'tcs'=>34,'tds_rate'=>35,'tds'=>36,
        'compensation'=>37,'claims'=>38,'recovery'=>39,
    ];
}

function detectPlatform($path, $forcedPlatform = 'auto')
{
    if ($forcedPlatform !== 'auto') return $forcedPlatform;

    $reader = new NativeXlsxReader();
    $reader->load($path);
    $sheetNames = $reader->getSheetNames();
    $reader->close();

    foreach ($sheetNames as $n) {
        if (stripos($n, 'order payments') !== false || stripos($n, 'ads cost') !== false) return 'meesho';
        if (stripos($n, 'referral payments') !== false) return 'meesho';
    }

    $nameLower = strtolower(basename($path));

    if (strpos($nameLower, 'meesho') !== false) return 'meesho';
    if (strpos($nameLower, 'flipkart') !== false || strpos($nameLower, 'fk') !== false) return 'flipkart';
    if (strpos($nameLower, 'amazon') !== false || strpos($nameLower, 'amz') !== false) return 'amazon';
    if (strpos($nameLower, 'myntra') !== false) return 'myntra';

    return 'generic';
}

function calculateIndianGST($summary)
{
    $rb = $summary['revenue_breakdown'];
    $taxableTurnover = $rb['net_product_revenue'];
    $avgGSTRate = 0.12;

    $outputGST  = $taxableTurnover * $avgGSTRate;
    $outputCGST = $outputGST / 2;
    $outputSGST = $outputGST / 2;
    $outputIGST = $outputGST;

    $gstOnAds          = abs($summary['ads_gst']);
    $gstOnCommission   = abs($summary['meesho_commission']) * 0.18 / 1.18;
    $gstOnShipping     = abs($summary['shipping_charge']) * 0.18 / 1.18;
    $gstOnWarehouse    = abs($summary['warehousing_fee']) * 0.18 / 1.18;
    $gstOnOtherSupport = abs($summary['gst_on_other']);
    $gstOnFixedFee     = abs($summary['fixed_fee']) * 0.18 / 1.18;
    $gstOnReturnShip   = abs($summary['return_shipping']) * 0.18 / 1.18;

    $totalITC = $gstOnAds + $gstOnCommission + $gstOnShipping + $gstOnWarehouse
              + $gstOnOtherSupport + $gstOnFixedFee + $gstOnReturnShip;

    $netGSTPayable = $outputGST - $totalITC;
    $tcs = abs($summary['tcs']);
    $tds = abs($summary['tds']);

    return [
        'taxable_turnover'   => $taxableTurnover,
        'avg_gst_rate'       => $avgGSTRate,
        'output_gst'         => $outputGST,
        'output_cgst'        => $outputCGST,
        'output_sgst'        => $outputSGST,
        'output_igst'        => $outputIGST,
        'itc_on_ads'         => $gstOnAds,
        'itc_on_commission'  => $gstOnCommission,
        'itc_on_shipping'    => $gstOnShipping,
        'itc_on_warehouse'   => $gstOnWarehouse,
        'itc_on_fixed_fee'   => $gstOnFixedFee,
        'itc_on_return_ship' => $gstOnReturnShip,
        'itc_on_other'       => $gstOnOtherSupport,
        'total_itc'          => $totalITC,
        'net_gst_payable'    => $netGSTPayable,
        'tcs'                => $tcs,
        'tds'                => $tds,
        'total_tax_paid'     => $netGSTPayable + $tcs + $tds,
    ];
}

// ================================================================
//  PROCESS MEESHO FILE
// ================================================================
function processMeeshoFile($path, &$orderDetails)
{
    $reader = new NativeXlsxReader();
    $reader->load($path);
    $sheets = $reader->getAllSheets();
    $reader->close();

    if (empty($sheets)) throw new Exception('No sheets found.');

    $r = [
        'platform'=>'Meesho','filename'=>basename($path),
        'sheet_names'=>array_keys($sheets),
        'total_rows'=>0,'total_orders'=>0,
        'sold'=>0,'sold_qty'=>0,'delivered'=>0,'delivered_qty'=>0,
        'shipped'=>0,'shipped_qty'=>0,'pending'=>0,'pending_qty'=>0,
        'return'=>0,'return_qty'=>0,'rto'=>0,'rto_qty'=>0,
        'cancelled'=>0,'cancelled_qty'=>0,'other'=>0,
        'listing_price'=>0,'total_sale_amount'=>0,'total_sale_return'=>0,
        'fixed_fee'=>0,'warehousing_fee'=>0,'return_premium'=>0,
        'meesho_commission'=>0,'meesho_gold_fee'=>0,'meesho_mall_fee'=>0,
        'return_shipping'=>0,'gst_compensation'=>0,'shipping_charge'=>0,
        'other_support'=>0,'waivers'=>0,'net_other_support'=>0,'gst_on_other'=>0,
        'tcs'=>0,'tds'=>0,'compensation'=>0,'claims'=>0,'recovery'=>0,
        'final_settlement'=>0,
        'ads_cost'=>0,'ads_cost_incl'=>0,'ads_gst'=>0,'ads_total'=>0,'ads_rows'=>0,
        'referral_net'=>0,'referral_tax'=>0,'referral_rows'=>0,
        'platform_compensation'=>0,'platform_comp_rows'=>0,
        'profit_calc'=>[
            'delivered_orders'=>0,'delivered_qty'=>0,
            'my_selling_total'=>0,'meesho_selling_total'=>0,'gross_profit_margin'=>0,
            'meesho_commission_cut'=>0,'meesho_other_fee_cut'=>0,'meesho_total_earned'=>0,
            'shipping_charged'=>0,'tax_deducted'=>0,'net_you_earned'=>0,
            'return_orders'=>0,'return_qty'=>0,'return_loss'=>0,'return_shipping_cut'=>0,
            'rto_orders'=>0,'rto_qty'=>0,'rto_loss'=>0,'rto_shipping_cut'=>0,
            'ads_cost'=>0,'ads_gst'=>0,'total_ads'=>0,
            'final_in_hand'=>0,'net_profit_loss'=>0,'per_unit_profit'=>0,
        ],
        'status_breakdown'=>[],'top_products'=>[],
        'payment_received'=>0,'ads_payment'=>0,'return_amount'=>0,
        'total_selling'=>0,'total_buy'=>0,'buy'=>0,'buy_qty'=>0,
        'total_commission'=>0,'total_shipping'=>0,'total_gst'=>0,
        'gross_profit'=>0,'net_profit'=>0,
        'state_breakdown'=>[],'top_states'=>[],
        'revenue_breakdown'=>[
            'gross_sales_pieces'=>0,'gross_sales_value'=>0,
            'rto_pieces'=>0,'rto_value'=>0,
            'customer_return_pieces'=>0,'customer_return_value'=>0,
            'net_sold_pieces'=>0,'net_product_revenue'=>0,
            'returned_stock_pieces'=>0,'returned_stock_value'=>0,
            'rto_courier_charges'=>0,'customer_return_shipping'=>0,
            'platform_fees_total'=>0,'final_cash_in_hand'=>0,
            'cost_price_total'=>0,'total_profit_loss'=>0,
        ],
    ];

    $orderSheet = null;
    foreach ($sheets as $name => $rows) {
        if (stripos($name, 'order') !== false && stripos($name, 'payment') !== false) {
            $orderSheet = $rows; break;
        }
    }

    if ($orderSheet) {
        $headerIdx = 0;
        foreach ($orderSheet as $i => $row) {
            foreach ($row as $cell) {
                if (stripos((string)$cell, 'sub order no') !== false) { $headerIdx = $i; break 2; }
            }
        }
        $dataRows = array_slice($orderSheet, $headerIdx + 1);
        if (!empty($dataRows)) {
            $first = $dataRows[0]; $isFormula = false;
            foreach ($first as $cell) {
                if (preg_match('/^[A-Z]\s*[=+\-]/', (string)$cell)) { $isFormula = true; break; }
            }
            if ($isFormula) array_shift($dataRows);
        }

        $C = meeshoCols();
        $uniqueOrders = [];

        foreach ($dataRows as $row) {
            $nonEmpty = array_filter($row, fn($v) => $v !== null && $v !== '');
            if (count($nonEmpty) < 3) continue;
            $orderNo = trim((string)($row[$C['sub_order_no']] ?? ''));
            if ($orderNo === '') continue;

            $r['total_rows']++;
            $uniqueOrders[$orderNo] = true;

            $statusRaw  = trim((string)($row[$C['status']] ?? ''));
            $statusType = detectStatus($statusRaw);
            $product = trim((string)($row[$C['product']] ?? '')) ?: 'Unknown';
            $sku     = trim((string)($row[$C['sku']] ?? ''));
            $qty     = max(1, (int)($row[$C['qty']] ?? 1));

            $myPrice      = cleanNumber($row[$C['listing_price']] ?? 0);
            $meeshoPrice  = cleanNumber($row[$C['total_sale_amount']] ?? 0);
            $commission   = cleanNumber($row[$C['meesho_commission']] ?? 0);
            $goldFee      = cleanNumber($row[$C['meesho_gold_fee']] ?? 0);
            $mallFee      = cleanNumber($row[$C['meesho_mall_fee']] ?? 0);
            $fixedFee     = cleanNumber($row[$C['fixed_fee_1']] ?? 0) + cleanNumber($row[$C['fixed_fee_2']] ?? 0);
            $warehouse    = cleanNumber($row[$C['warehousing_fee_1']] ?? 0) + cleanNumber($row[$C['warehousing_fee_2']] ?? 0);
            $shipping     = cleanNumber($row[$C['shipping_charge']] ?? 0);
            $retShipping  = cleanNumber($row[$C['return_shipping']] ?? 0);
            $retPremium   = cleanNumber($row[$C['return_premium']] ?? 0) + cleanNumber($row[$C['return_premium_return']] ?? 0);
            $otherSupport = cleanNumber($row[$C['other_support']] ?? 0);
            $netOther     = cleanNumber($row[$C['net_other_support']] ?? 0);
            $gstOnOther   = cleanNumber($row[$C['gst_on_other']] ?? 0);
            $tcs          = cleanNumber($row[$C['tcs']] ?? 0);
            $tds          = cleanNumber($row[$C['tds']] ?? 0);
            $finalSettle  = cleanNumber($row[$C['final_settlement']] ?? 0);
            $saleReturn   = cleanNumber($row[$C['total_sale_return']] ?? 0);

            switch ($statusType) {
                case 'delivered': $r['delivered']++; $r['delivered_qty']+=$qty; $r['sold']++; $r['sold_qty']+=$qty; break;
                case 'shipped':   $r['shipped']++;   $r['shipped_qty']+=$qty;   $r['sold']++; $r['sold_qty']+=$qty; break;
                case 'pending':   $r['pending']++;   $r['pending_qty']+=$qty;   break;
                case 'cancelled': $r['cancelled']++; $r['cancelled_qty']+=$qty; break;
                case 'return':    $r['return']++;    $r['return_qty']+=$qty;    break;
                case 'rto':       $r['rto']++;       $r['rto_qty']+=$qty;       break;
                default:          $r['other']++;
            }

            $r['listing_price']     += $myPrice;
            $r['total_sale_amount'] += $meeshoPrice;
            $r['total_sale_return'] += $saleReturn;
            $r['fixed_fee']         += $fixedFee;
            $r['warehousing_fee']   += $warehouse;
            $r['return_premium']    += $retPremium;
            $r['meesho_commission'] += $commission;
            $r['meesho_gold_fee']   += $goldFee;
            $r['meesho_mall_fee']   += $mallFee;
            $r['return_shipping']   += $retShipping;
            $r['gst_compensation']  += cleanNumber($row[$C['gst_compensation']] ?? 0);
            $r['shipping_charge']   += $shipping;
            $r['other_support']     += $otherSupport;
            $r['waivers']           += cleanNumber($row[$C['waivers']] ?? 0);
            $r['net_other_support'] += $netOther;
            $r['gst_on_other']      += $gstOnOther;
            $r['tcs']               += $tcs;
            $r['tds']               += $tds;
            $r['compensation']      += cleanNumber($row[$C['compensation']] ?? 0);
            $r['claims']            += cleanNumber($row[$C['claims']] ?? 0);
            $r['recovery']          += cleanNumber($row[$C['recovery']] ?? 0);
            $r['final_settlement']  += $finalSettle;

            $rb = &$r['revenue_breakdown'];
            $rb['gross_sales_pieces'] += $qty;
            $rb['gross_sales_value']  += $myPrice * $qty;

            if ($statusType === 'rto') {
                $rb['rto_pieces'] += $qty;
                $rb['rto_value']  += $myPrice * $qty;
                $rb['rto_courier_charges'] += abs($retShipping);
                $rb['returned_stock_pieces'] += $qty;
                $rb['returned_stock_value']  += $myPrice * $qty;
            } elseif ($statusType === 'return') {
                $rb['customer_return_pieces'] += $qty;
                $rb['customer_return_value']  += $myPrice * $qty;
                $rb['customer_return_shipping'] += abs($retShipping);
                $rb['returned_stock_pieces'] += $qty;
                $rb['returned_stock_value']  += $myPrice * $qty;
            } elseif ($statusType === 'delivered') {
                $rb['net_sold_pieces'] += $qty;
                $rb['net_product_revenue'] += $myPrice * $qty;
                $rb['platform_fees_total'] += abs($commission) + abs($goldFee) + abs($mallFee)
                                           + abs($fixedFee) + abs($warehouse)
                                           + abs($shipping) + abs($otherSupport) + abs($netOther);
                $rb['final_cash_in_hand'] += $finalSettle;
            }

            $pc = &$r['profit_calc'];
            if ($statusType === 'delivered') {
                $pc['delivered_orders']++;
                $pc['delivered_qty'] += $qty;
                $pc['my_selling_total']     += $myPrice * $qty;
                $pc['meesho_selling_total'] += $meeshoPrice * $qty;
                $pc['gross_profit_margin']  += ($meeshoPrice - $myPrice) * $qty;
                $meeshoCut = abs($commission) + abs($goldFee) + abs($mallFee);
                $otherFee  = abs($fixedFee) + abs($warehouse) + abs($otherSupport) + abs($netOther);
                $pc['meesho_commission_cut'] += $meeshoCut;
                $pc['meesho_other_fee_cut']  += $otherFee;
                $pc['meesho_total_earned']   += $meeshoCut + $otherFee;
                $pc['shipping_charged']      += abs($shipping);
                $pc['tax_deducted']          += abs($tcs) + abs($tds) + abs($gstOnOther);
                $pc['net_you_earned']        += $finalSettle;
            } elseif ($statusType === 'return') {
                $pc['return_orders']++; $pc['return_qty'] += $qty;
                $pc['return_loss'] += abs($saleReturn);
                $pc['return_shipping_cut'] += abs($retShipping);
            } elseif ($statusType === 'rto') {
                $pc['rto_orders']++; $pc['rto_qty'] += $qty;
                $pc['rto_loss'] += abs($saleReturn);
                $pc['rto_shipping_cut'] += abs($retShipping);
            }

            $sk = $statusRaw ?: 'Unknown';
            $r['status_breakdown'][$sk] = ($r['status_breakdown'][$sk] ?? 0) + 1;
            $r['top_products'][$product] = ($r['top_products'][$product] ?? 0) + $qty;

            $orderDetails[] = [
                'order_no'=>$orderNo,'product'=>$product,'sku'=>$sku,'qty'=>$qty,
                'status_raw'=>$statusRaw,'status_type'=>$statusType,
                'my_price'=>$myPrice,'meesho_price'=>$meeshoPrice,
                'total_sale_return'=>$saleReturn,
                'my_total'=>$myPrice*$qty,'meesho_total'=>$meeshoPrice*$qty,
                'margin'=>($meeshoPrice - $myPrice) * $qty,
                'commission'=>$commission,'gold_fee'=>$goldFee,'mall_fee'=>$mallFee,
                'fixed_fee'=>$fixedFee,'warehousing_fee'=>$warehouse,
                'shipping_charge'=>$shipping,'return_shipping'=>$retShipping,
                'return_premium'=>$retPremium,'other_support'=>$otherSupport,
                'net_other'=>$netOther,'gst_on_other'=>$gstOnOther,
                'tcs'=>$tcs,'tds'=>$tds,'final_settlement'=>$finalSettle,
            ];
        }
        $r['total_orders'] = count($uniqueOrders);
    }

    foreach ($sheets as $name => $srows) {
        if (stripos($name, 'ads') !== false) {
            $hIdx = 0;
            foreach ($srows as $i => $row) {
                foreach ($row as $c) {
                    $cs = strtolower(trim((string)$c));
                    if ($cs === 'deduction date' || $cs === 'ad cost') { $hIdx = $i; break 2; }
                }
            }
            foreach (array_slice($srows, $hIdx + 1) as $row) {
                $ne = array_filter($row, fn($v) => $v !== null && $v !== '');
                if (count($ne) < 2) continue;
                $first = trim((string)($row[0] ?? ''));
                if ($first === '' || preg_match('/^[A-Z]\s*[=+\-]/', $first)) continue;
                $r['ads_rows']++;
                $r['ads_cost']      += abs(cleanNumber($row[3] ?? 0));
                $r['ads_cost_incl'] += abs(cleanNumber($row[5] ?? 0));
                $r['ads_gst']       += abs(cleanNumber($row[6] ?? 0));
                $r['ads_total']     += abs(cleanNumber($row[7] ?? 0));
            }
            break;
        }
    }

    foreach ($sheets as $name => $rows) {
        if (stripos($name, 'referral') !== false) {
            foreach ($rows as $row) {
                if (empty($row)) continue;
                $first = strtolower((string)($row[0] ?? ''));
                if (strpos($first, 'no data') !== false) break;
                if (strpos($first, 'reward id') !== false) continue;
                if (strpos($first, 'referral payment') !== false) continue;
                $net = cleanNumber($row[4] ?? 0); $tax = cleanNumber($row[5] ?? 0);
                if ($net != 0 || $tax != 0) {
                    $r['referral_rows']++; $r['referral_net'] += $net; $r['referral_tax'] += $tax;
                }
            }
            break;
        }
    }

    foreach ($sheets as $name => $rows) {
        if (stripos($name, 'compensation') !== false) {
            foreach ($rows as $row) {
                if (empty($row)) continue;
                $first = strtolower((string)($row[0] ?? ''));
                if (strpos($first, 'no data') !== false) break;
                if (strpos($first, 'date') !== false) continue;
                if (strpos($first, 'platform recovery') !== false) continue;
                $amt = cleanNumber($row[3] ?? 0);
                if ($amt != 0) { $r['platform_comp_rows']++; $r['platform_compensation'] += $amt; }
            }
            break;
        }
    }

    arsort($r['top_products']);
    $r['top_products'] = array_slice($r['top_products'], 0, 10, true);

    $rb = &$r['revenue_breakdown'];
    $rb['cost_price_total'] = 0;

    $pc = &$r['profit_calc'];
    $pc['ads_cost'] = $r['ads_cost'];
    $pc['ads_gst']  = $r['ads_gst'];
    $pc['total_ads'] = $r['ads_total'];
    $pc['final_in_hand'] = $r['final_settlement'];
    $pc['net_profit_loss'] = $r['final_settlement'] - $r['ads_total'];
    if ($pc['delivered_qty'] > 0) $pc['per_unit_profit'] = $pc['net_profit_loss'] / $pc['delivered_qty'];

    $r['total_commission'] = $r['meesho_commission'] + $r['meesho_gold_fee'] + $r['meesho_mall_fee'];
    $r['total_shipping']   = $r['shipping_charge'] + $r['return_shipping'];
    $r['total_gst']        = $r['tcs'] + $r['tds'] + $r['gst_on_other'];
    $r['total_selling']    = $r['total_sale_amount'];
    $r['payment_received'] = $r['final_settlement'];
    $r['ads_payment']      = $r['ads_total'];
    $r['return_amount']    = abs($r['total_sale_return']);
    $r['gross_profit']     = $r['total_selling'];
    $r['net_profit']       = $r['final_settlement'] - $r['ads_total'];

    return $r;
}

// ================================================================
//  PROCESS FLIPKART FILE
// ================================================================
function processFlipkartFile($path, &$orderDetails)
{
    $reader = new NativeXlsxReader();
    $reader->load($path);
    $sheets = $reader->getAllSheets();
    $reader->close();

    if (empty($sheets)) throw new Exception('No sheets found.');

    $dataRows = []; $bestCount = 0;
    foreach ($sheets as $name => $rows) {
        $neCount = 0;
        foreach ($rows as $r) {
            if (count(array_filter($r, fn($v) => $v !== null && $v !== '')) >= 3) $neCount++;
        }
        if ($neCount > $bestCount) { $bestCount = $neCount; $dataRows = $rows; }
    }

    if (empty($dataRows)) throw new Exception('No data found.');

    $headerIdx = 0;
    foreach ($dataRows as $i => $row) {
        foreach ($row as $cell) {
            $c = strtolower(trim((string)$cell));
            if (strpos($c, 'order item id') !== false || $c === 'order id' || strpos($c, 'order id') !== false) {
                $headerIdx = $i; break 2;
            }
        }
    }
    if ($headerIdx === 0) {
        foreach ($dataRows as $i => $row) {
            if (count(array_filter($row, fn($v) => $v !== null && $v !== '')) > 5) { $headerIdx = $i; break; }
        }
    }

    $headers  = $dataRows[$headerIdx];
    $bodyRows = array_slice($dataRows, $headerIdx + 1);

    if (!empty($bodyRows)) {
        $first = $bodyRows[0]; $isLabel = false;
        foreach ($first as $cell) {
            if (preg_match('/^[A-Z]\s*[=+\-]/', (string)$cell)) { $isLabel = true; break; }
        }
        if ($isLabel) array_shift($bodyRows);
    }

    $norm = function($k) { return strtolower(trim(preg_replace('/[^a-z0-9]+/i', '_', (string)$k))); };
    $find = function($names) use ($headers, $norm) {
        $nh = array_map($norm, $headers);
        foreach ($names as $n) {
            $idx = array_search($norm($n), $nh);
            if ($idx !== false) return $idx;
        }
        return false;
    };

    $C = [
        'order_id'          => $find(['order item id','order id','order_id','order no']),
        'product'           => $find(['product name','product','item name','product description','sku']),
        'sku'               => $find(['sku','seller sku','seller_sku']),
        'qty'               => $find(['quantity','qty','units']),
        'status'            => $find(['order status','live order status','status','shipment status']),
        'selling_price'     => $find(['selling price','final selling price','supplier price','item price','total price','order item value']),
        'marketplace_fee'   => $find(['marketplace fee','marketplace_fee','marketplace fees']),
        'commission'        => $find(['commission','commission fee','category commission']),
        'fixed_fee'         => $find(['fixed fee','fixed_fee']),
        'collection_fee'    => $find(['collection fee','collection_fee']),
        'shipping_fee'      => $find(['shipping fee','shipping_fee','shipping charge','forward shipping']),
        'return_shipping'   => $find(['return shipping fee','return shipping charge','reverse shipping']),
        'return_premium'    => $find(['return premium','return_premium']),
        'gst'               => $find(['gst','tax','gst amount','total gst']),
        'tcs'               => $find(['tcs','tcs amount','tax collected']),
        'tds'               => $find(['tds','tds amount','tax deducted']),
        'net_settlement'    => $find(['net settlement','net_settlement','net amount','final settlement','payout','bank settlement','neft amount']),
        'ads_cost'          => $find(['ads cost','ad spend','pla cost','flipkart ads']),
    ];

    $r = [
        'platform'=>'Flipkart','filename'=>basename($path),
        'sheet_names'=>array_keys($sheets),
        'total_rows'=>0,'total_orders'=>0,
        'sold'=>0,'sold_qty'=>0,'delivered'=>0,'delivered_qty'=>0,
        'shipped'=>0,'shipped_qty'=>0,'pending'=>0,'pending_qty'=>0,
        'return'=>0,'return_qty'=>0,'rto'=>0,'rto_qty'=>0,
        'cancelled'=>0,'cancelled_qty'=>0,'other'=>0,
        'listing_price'=>0,'total_sale_amount'=>0,'total_sale_return'=>0,
        'marketplace_fee'=>0,'commission'=>0,'fixed_fee'=>0,'collection_fee'=>0,
        'shipping_fee'=>0,'return_shipping'=>0,'return_premium'=>0,
        'gst'=>0,'tcs'=>0,'tds'=>0,
        'final_settlement'=>0,'ads_total'=>0,'ads_rows'=>0,
        'meesho_commission'=>0,'meesho_gold_fee'=>0,'meesho_mall_fee'=>0,
        'warehousing_fee'=>0,'other_support'=>0,'waivers'=>0,'net_other_support'=>0,
        'gst_on_other'=>0,'shipping_charge'=>0,'compensation'=>0,'claims'=>0,'recovery'=>0,
        'referral_net'=>0,'referral_tax'=>0,'referral_rows'=>0,
        'platform_compensation'=>0,'platform_comp_rows'=>0,
        'ads_cost'=>0,'ads_cost_incl'=>0,'ads_gst'=>0,
        'profit_calc'=>[
            'delivered_orders'=>0,'delivered_qty'=>0,
            'my_selling_total'=>0,'meesho_selling_total'=>0,'gross_profit_margin'=>0,
            'meesho_commission_cut'=>0,'meesho_other_fee_cut'=>0,'meesho_total_earned'=>0,
            'shipping_charged'=>0,'tax_deducted'=>0,'net_you_earned'=>0,
            'return_orders'=>0,'return_qty'=>0,'return_loss'=>0,'return_shipping_cut'=>0,
            'rto_orders'=>0,'rto_qty'=>0,'rto_loss'=>0,'rto_shipping_cut'=>0,
            'ads_cost'=>0,'ads_gst'=>0,'total_ads'=>0,
            'final_in_hand'=>0,'net_profit_loss'=>0,'per_unit_profit'=>0,
        ],
        'status_breakdown'=>[],'top_products'=>[],
        'payment_received'=>0,'ads_payment'=>0,'return_amount'=>0,
        'total_selling'=>0,'total_buy'=>0,'buy'=>0,'buy_qty'=>0,
        'total_commission'=>0,'total_shipping'=>0,'total_gst'=>0,
        'gross_profit'=>0,'net_profit'=>0,
        'state_breakdown'=>[],'top_states'=>[],
        'revenue_breakdown'=>[
            'gross_sales_pieces'=>0,'gross_sales_value'=>0,
            'rto_pieces'=>0,'rto_value'=>0,
            'customer_return_pieces'=>0,'customer_return_value'=>0,
            'net_sold_pieces'=>0,'net_product_revenue'=>0,
            'returned_stock_pieces'=>0,'returned_stock_value'=>0,
            'rto_courier_charges'=>0,'customer_return_shipping'=>0,
            'platform_fees_total'=>0,'final_cash_in_hand'=>0,
            'cost_price_total'=>0,'total_profit_loss'=>0,
        ],
    ];

    $uniqueOrders = [];

    foreach ($bodyRows as $row) {
        $nonEmpty = array_filter($row, fn($v) => $v !== null && $v !== '');
        if (count($nonEmpty) < 3) continue;

        $get = function($key) use ($row, $C) {
            $idx = $C[$key] ?? false;
            if ($idx === false || !isset($row[$idx])) return '';
            return trim((string)$row[$idx]);
        };

        $orderNo = $get('order_id');
        if ($orderNo === '') continue;

        $r['total_rows']++;
        $uniqueOrders[$orderNo] = true;

        $statusRaw  = $get('status');
        $statusType = detectStatus($statusRaw);
        $product = $get('product') ?: ($get('sku') ?: 'Unknown');
        $sku     = $get('sku');
        $qty     = max(1, (int)($get('qty') ?: 1));

        $selling       = cleanNumber($get('selling_price'));
        $mkFee         = cleanNumber($get('marketplace_fee'));
        $commission    = cleanNumber($get('commission'));
        $fixedFee      = cleanNumber($get('fixed_fee'));
        $collectionFee = cleanNumber($get('collection_fee'));
        $shipping      = cleanNumber($get('shipping_fee'));
        $retShipping   = cleanNumber($get('return_shipping'));
        $retPremium    = cleanNumber($get('return_premium'));
        $gst           = cleanNumber($get('gst'));
        $tcs           = cleanNumber($get('tcs'));
        $tds           = cleanNumber($get('tds'));
        $netSettle     = cleanNumber($get('net_settlement'));
        $ads           = cleanNumber($get('ads_cost'));

        switch ($statusType) {
            case 'delivered': $r['delivered']++; $r['delivered_qty'] += $qty; $r['sold']++; $r['sold_qty'] += $qty; break;
            case 'shipped':   $r['shipped']++;   $r['shipped_qty'] += $qty;   $r['sold']++; $r['sold_qty'] += $qty; break;
            case 'pending':   $r['pending']++;   $r['pending_qty'] += $qty;   break;
            case 'cancelled': $r['cancelled']++; $r['cancelled_qty'] += $qty; break;
            case 'return':    $r['return']++;    $r['return_qty'] += $qty;    break;
            case 'rto':       $r['rto']++;       $r['rto_qty'] += $qty;       break;
            default:          $r['other']++;
        }

        $r['listing_price']     += $selling;
        $r['marketplace_fee']   += abs($mkFee);
        $r['commission']        += abs($commission);
        $r['fixed_fee']         += abs($fixedFee);
        $r['collection_fee']    += abs($collectionFee);
        $r['shipping_fee']      += abs($shipping);
        $r['return_shipping']   += abs($retShipping);
        $r['return_premium']    += abs($retPremium);
        $r['gst']               += abs($gst);
        $r['tcs']               += abs($tcs);
        $r['tds']               += abs($tds);
        $r['final_settlement']  += $netSettle;
        $r['ads_total']         += abs($ads);
        if ($ads != 0) $r['ads_rows']++;

        $rb = &$r['revenue_breakdown'];
        $rb['gross_sales_pieces'] += $qty;
        $rb['gross_sales_value']  += $selling * $qty;

        if ($statusType === 'rto') {
            $rb['rto_pieces'] += $qty;
            $rb['rto_value']  += $selling * $qty;
            $rb['rto_courier_charges'] += abs($retShipping);
            $rb['returned_stock_pieces'] += $qty;
            $rb['returned_stock_value']  += $selling * $qty;
        } elseif ($statusType === 'return') {
            $rb['customer_return_pieces'] += $qty;
            $rb['customer_return_value']  += $selling * $qty;
            $rb['customer_return_shipping'] += abs($retShipping);
            $rb['returned_stock_pieces'] += $qty;
            $rb['returned_stock_value']  += $selling * $qty;
        } elseif ($statusType === 'delivered') {
            $rb['net_sold_pieces'] += $qty;
            $rb['net_product_revenue'] += $selling * $qty;
            $rb['platform_fees_total'] += abs($mkFee) + abs($commission) + abs($fixedFee)
                                       + abs($collectionFee) + abs($shipping);
            $rb['final_cash_in_hand'] += $netSettle;
        }

        $pc = &$r['profit_calc'];
        if ($statusType === 'delivered') {
            $pc['delivered_orders']++;
            $pc['delivered_qty'] += $qty;
            $pc['my_selling_total']     += $selling * $qty;
            $pc['gross_profit_margin']  += $selling * $qty;
            $pc['fees_cut']             += abs($mkFee) + abs($commission) + abs($fixedFee)
                                         + abs($collectionFee);
            $pc['shipping_charged']     += abs($shipping);
            $pc['tax_deducted']         += abs($gst) + abs($tcs) + abs($tds);
            $pc['net_you_earned']       += $netSettle;
        } elseif ($statusType === 'return') {
            $pc['return_orders']++; $pc['return_qty'] += $qty;
            $pc['return_loss'] += $selling * $qty;
            $pc['return_shipping_cut'] += abs($retShipping);
        } elseif ($statusType === 'rto') {
            $pc['rto_orders']++; $pc['rto_qty'] += $qty;
            $pc['rto_loss'] += $selling * $qty;
            $pc['rto_shipping_cut'] += abs($retShipping);
        }

        $sk = $statusRaw ?: 'Unknown';
        $r['status_breakdown'][$sk] = ($r['status_breakdown'][$sk] ?? 0) + 1;
        $r['top_products'][$product] = ($r['top_products'][$product] ?? 0) + $qty;

        $orderDetails[] = [
            'order_no'=>$orderNo,'product'=>$product,'sku'=>$sku,'qty'=>$qty,
            'status_raw'=>$statusRaw,'status_type'=>$statusType,
            'my_price'=>$selling,'meesho_price'=>$selling,
            'total_sale_return'=>abs($retShipping),
            'my_total'=>$selling*$qty,'meesho_total'=>$selling*$qty,
            'margin'=>0,
            'commission'=>$commission,'gold_fee'=>0,'mall_fee'=>0,
            'fixed_fee'=>$fixedFee,'warehousing_fee'=>$collectionFee,
            'shipping_charge'=>$shipping,'return_shipping'=>$retShipping,
            'return_premium'=>$retPremium,'other_support'=>$mkFee,
            'net_other'=>0,'gst_on_other'=>$gst,
            'tcs'=>$tcs,'tds'=>$tds,'final_settlement'=>$netSettle,
        ];
    }

    $r['total_orders'] = count($uniqueOrders);
    arsort($r['top_products']);
    $r['top_products'] = array_slice($r['top_products'], 0, 10, true);

    $rb = &$r['revenue_breakdown'];
    $rb['cost_price_total'] = 0;

    $pc = &$r['profit_calc'];
    $pc['total_ads'] = $r['ads_total'];
    $pc['final_in_hand'] = $r['final_settlement'];

    $netSales = $rb['net_product_revenue'];
    $adsCost = abs($r['ads_total']);
    $finalCash = $netSales - $adsCost;
    $returnCharges = $rb['rto_courier_charges'] + $rb['customer_return_shipping'];
    $platformFees  = $rb['platform_fees_total'];
    $rb['final_cash_in_hand'] = $finalCash;
    $rb['total_profit_loss'] = $finalCash - $returnCharges - $platformFees - $rb['cost_price_total'];

    $pc['net_profit_loss'] = $rb['total_profit_loss'];
    if ($pc['delivered_qty'] > 0) $pc['per_unit_profit'] = $pc['net_profit_loss'] / $pc['delivered_qty'];

    $r['total_commission'] = $r['commission'] + $r['fixed_fee'] + $r['collection_fee'];
    $r['total_shipping']   = $r['shipping_fee'] + $r['return_shipping'];
    $r['total_gst']        = $r['gst'] + $r['tcs'] + $r['tds'];
    $r['payment_received'] = $r['final_settlement'];
    $r['ads_payment']      = $r['ads_total'];

    return $r;
}

// ================================================================
//  GENERIC FILE
// ================================================================
function analyzeGenericFile($path, $originalName, &$orderDetails, $platform = 'Other')
{
    $reader = new NativeXlsxReader();
    $reader->load($path);
    $rows = $reader->getRows();
    $reader->close();

    if (count($rows) < 2) throw new Exception('File has no data rows.');

    $headerIdx = 0;
    foreach ($rows as $i => $row) {
        if (count(array_filter($row, fn($v) => $v !== null && $v !== '')) > 3) { $headerIdx = $i; break; }
    }
    $headers  = $rows[$headerIdx];
    $dataRows = array_slice($rows, $headerIdx + 1);

    $norm = function($k) { return strtolower(trim(preg_replace('/[^a-z0-9]+/i', '_', (string)$k))); };
    $find = function($names) use ($headers, $norm) {
        $nh = array_map($norm, $headers);
        foreach ($names as $n) {
            $idx = array_search($norm($n), $nh);
            if ($idx !== false) return $idx;
        }
        return false;
    };

    $C = [
        'order_id'      => $find(['order id','order_id','sub order no','order no','order item id','amazon-order-id']),
        'product'       => $find(['product name','product','item name','sku','seller sku']),
        'status'        => $find(['order status','live order status','status','shipment status']),
        'qty'           => $find(['quantity','qty','units']),
        'selling_price' => $find(['selling price','supplier price','listing price','total price','item price','final price']),
        'payment'       => $find(['payment','settlement','net amount','final settlement','net payment','payout','amount credited']),
        'ads'           => $find(['ads','ads cost','ad spend','marketing','advertising']),
        'return_amount' => $find(['return amount','refund amount','return charges']),
        'commission'    => $find(['commission','marketplace fee','platform fee']),
        'shipping'      => $find(['shipping charge','shipping','delivery charge']),
        'gst'           => $find(['gst','tax','tcs','tds']),
    ];

    $s = array_fill_keys([
        'total_rows','total_orders','sold','sold_qty','delivered','delivered_qty',
        'shipped','shipped_qty','pending','pending_qty','return','return_qty',
        'rto','rto_qty','cancelled','cancelled_qty','other',
        'payment_received','ads_payment','return_amount','total_selling','total_buy',
        'buy','buy_qty','total_commission','total_shipping','total_gst',
        'gross_profit','net_profit','listing_price','total_sale_amount','total_sale_return',
        'meesho_commission','meesho_gold_fee','meesho_mall_fee','fixed_fee','warehousing_fee',
        'return_shipping','return_premium','shipping_charge','gst_compensation','other_support',
        'waivers','net_other_support','gst_on_other','tcs','tds',
        'compensation','claims','recovery','final_settlement',
        'ads_cost','ads_cost_incl','ads_gst','ads_total','ads_rows',
        'referral_net','referral_tax','referral_rows','platform_compensation','platform_comp_rows'
    ], 0);

    $s['platform'] = $platform;
    $s['filename'] = $originalName;
    $s['sheet_names'] = ['(single sheet)'];
    $s['status_breakdown'] = []; $s['top_products'] = [];
    $s['state_breakdown'] = []; $s['top_states'] = [];

    $s['profit_calc'] = array_fill_keys([
        'delivered_orders','delivered_qty','my_selling_total','meesho_selling_total',
        'gross_profit_margin','meesho_commission_cut','meesho_other_fee_cut',
        'meesho_total_earned','shipping_charged','tax_deducted','net_you_earned',
        'return_orders','return_qty','return_loss','return_shipping_cut',
        'rto_orders','rto_qty','rto_loss','rto_shipping_cut',
        'ads_cost','ads_gst','total_ads','final_in_hand','net_profit_loss','per_unit_profit',
    ], 0);

    $s['revenue_breakdown'] = array_fill_keys([
        'gross_sales_pieces','gross_sales_value','rto_pieces','rto_value',
        'customer_return_pieces','customer_return_value',
        'net_sold_pieces','net_product_revenue',
        'returned_stock_pieces','returned_stock_value',
        'rto_courier_charges','customer_return_shipping',
        'platform_fees_total','final_cash_in_hand',
        'cost_price_total','total_profit_loss'
    ], 0);

    $uniqueOrders = [];
    foreach ($dataRows as $row) {
        $ne = array_filter($row, fn($v) => $v !== null && $v !== '');
        if (count($ne) < 2) continue;
        $s['total_rows']++;

        $g = function($k) use ($row, $C) {
            $idx = $C[$k] ?? false;
            if ($idx === false || !isset($row[$idx])) return '';
            return trim((string)$row[$idx]);
        };

        $oid = $g('order_id'); if ($oid !== '') $uniqueOrders[$oid] = true;
        $statusType = detectStatus($g('status'));
        $statusRaw = $g('status');
        $product = $g('product') ?: 'Unknown';
        $qty = max(1, (int)($g('qty') ?: 1));

        $selling = cleanNumber($g('selling_price'));
        $payment = cleanNumber($g('payment'));
        $ads     = cleanNumber($g('ads'));
        $retAmt  = cleanNumber($g('return_amount'));
        $commission = cleanNumber($g('commission'));
        $shipping = cleanNumber($g('shipping'));
        $gst     = cleanNumber($g('gst'));

        switch ($statusType) {
            case 'delivered': $s['delivered']++; $s['delivered_qty']+=$qty; $s['sold']++; $s['sold_qty']+=$qty; break;
            case 'shipped':   $s['shipped']++;   $s['shipped_qty']+=$qty;   $s['sold']++; $s['sold_qty']+=$qty; break;
            case 'pending':   $s['pending']++;   $s['pending_qty']+=$qty;   break;
            case 'cancelled': $s['cancelled']++; $s['cancelled_qty']+=$qty; break;
            case 'return':    $s['return']++;    $s['return_qty']+=$qty;    break;
            case 'rto':       $s['rto']++;       $s['rto_qty']+=$qty;       break;
            default:          $s['other']++;
        }

        $s['total_selling'] += $selling;
        $s['payment_received'] += $payment;
        $s['ads_payment'] += $ads;
        $s['return_amount'] += $retAmt;
        $s['total_commission'] += $commission;
        $s['total_shipping'] += $shipping;
        $s['total_gst'] += $gst;

        $pc = &$s['profit_calc'];
        if ($statusType === 'delivered') {
            $pc['delivered_orders']++; $pc['delivered_qty'] += $qty;
            $pc['my_selling_total'] += $selling * $qty;
            $pc['meesho_selling_total'] += $selling * $qty;
            $pc['meesho_commission_cut'] += abs($commission);
            $pc['meesho_other_fee_cut'] += abs($shipping);
            $pc['meesho_total_earned'] += abs($commission) + abs($shipping);
            $pc['shipping_charged'] += abs($shipping);
            $pc['tax_deducted'] += abs($gst);
            $pc['net_you_earned'] += $payment;
        } elseif ($statusType === 'return') {
            $pc['return_orders']++; $pc['return_qty'] += $qty;
            $pc['return_loss'] += abs($retAmt);
        } elseif ($statusType === 'rto') {
            $pc['rto_orders']++; $pc['rto_qty'] += $qty;
            $pc['rto_loss'] += abs($retAmt);
        }

        $rb = &$s['revenue_breakdown'];
        $rb['gross_sales_pieces'] += $qty;
        $rb['gross_sales_value']  += $selling * $qty;
        if ($statusType === 'rto') {
            $rb['rto_pieces'] += $qty; $rb['rto_value'] += $selling * $qty;
            $rb['returned_stock_pieces'] += $qty; $rb['returned_stock_value'] += $selling * $qty;
        } elseif ($statusType === 'return') {
            $rb['customer_return_pieces'] += $qty; $rb['customer_return_value'] += $selling * $qty;
            $rb['returned_stock_pieces'] += $qty; $rb['returned_stock_value'] += $selling * $qty;
        } elseif ($statusType === 'delivered') {
            $rb['net_sold_pieces'] += $qty;
            $rb['net_product_revenue'] += $selling * $qty;
            $rb['platform_fees_total'] += abs($commission) + abs($shipping) + abs($gst);
            $rb['final_cash_in_hand'] += $payment;
        }

        $sk = $statusRaw ?: 'Unknown';
        $s['status_breakdown'][$sk] = ($s['status_breakdown'][$sk] ?? 0) + 1;
        $s['top_products'][$product] = ($s['top_products'][$product] ?? 0) + $qty;

        $orderDetails[] = [
            'order_no'=>$oid,'product'=>$product,'sku'=>'','qty'=>$qty,
            'status_raw'=>$statusRaw,'status_type'=>$statusType,
            'my_price'=>$selling,'meesho_price'=>$selling,
            'total_sale_return'=>$retAmt,
            'my_total'=>$selling*$qty,'meesho_total'=>$selling*$qty,
            'margin'=>0,
            'commission'=>$commission,'gold_fee'=>0,'mall_fee'=>0,
            'fixed_fee'=>0,'warehousing_fee'=>0,
            'shipping_charge'=>$shipping,'return_shipping'=>$retAmt,
            'return_premium'=>0,'other_support'=>0,'net_other'=>0,
            'gst_on_other'=>$gst,'tcs'=>0,'tds'=>0,
            'final_settlement'=>$payment,
        ];
    }

    $s['total_orders'] = count($uniqueOrders);
    arsort($s['top_products']);
    $s['top_products'] = array_slice($s['top_products'], 0, 10, true);

    $s['gross_profit'] = $s['total_selling'] - $s['total_buy'];
    $s['net_profit'] = $s['total_selling'] - $s['ads_payment'] - $s['total_commission']
                     - $s['total_shipping'] - $s['return_amount'];

    $pc = &$s['profit_calc'];
    $pc['total_ads'] = $s['ads_payment'];
    $pc['final_in_hand'] = $s['payment_received'];
    $pc['net_profit_loss'] = $s['payment_received'] - $s['ads_payment'];
    if ($pc['delivered_qty'] > 0) $pc['per_unit_profit'] = $pc['net_profit_loss'] / $pc['delivered_qty'];

    return $s;
}

// ========== DISPATCHER ==========
function analyzeFile($path, $originalName, &$orderDetails, $forcedPlatform = 'auto')
{
    $detected = detectPlatform($path, $forcedPlatform);

    switch ($detected) {
        case 'meesho':   return processMeeshoFile($path, $orderDetails);
        case 'flipkart': return processFlipkartFile($path, $orderDetails);
        case 'amazon':   return analyzeGenericFile($path, $originalName, $orderDetails, 'Amazon');
        case 'myntra':   return analyzeGenericFile($path, $originalName, $orderDetails, 'Myntra');
        default:
            $ext = strtolower(pathinfo($path, PATHINFO_EXTENSION));
            if ($ext === 'xlsx') {
                $reader = new NativeXlsxReader();
                $reader->load($path);
                $names = $reader->getSheetNames();
                $reader->close();
                foreach ($names as $n) {
                    if (stripos($n, 'order payments') !== false || stripos($n, 'ads cost') !== false) {
                        return processMeeshoFile($path, $orderDetails);
                    }
                }
            }
            return analyzeGenericFile($path, $originalName, $orderDetails, 'Other');
    }
}

function mergeResults($results)
{
    $keys = [
        'total_rows','total_orders','sold','sold_qty','delivered','delivered_qty',
        'shipped','shipped_qty','pending','pending_qty','return','return_qty',
        'rto','rto_qty','cancelled','cancelled_qty','other',
        'payment_received','ads_payment','return_amount','total_selling','total_buy',
        'buy','buy_qty','total_commission','total_shipping','total_gst',
        'gross_profit','net_profit','listing_price','total_sale_amount','total_sale_return',
        'meesho_commission','meesho_gold_fee','meesho_mall_fee','fixed_fee','warehousing_fee',
        'return_shipping','return_premium','shipping_charge','gst_compensation','other_support',
        'waivers','net_other_support','gst_on_other','tcs','tds',
        'compensation','claims','recovery','final_settlement',
        'ads_cost','ads_cost_incl','ads_gst','ads_total','ads_rows',
        'referral_net','referral_tax','referral_rows','platform_compensation','platform_comp_rows',
        'marketplace_fee','commission','collection_fee','shipping_fee','gst'
    ];
    $merged = array_fill_keys($keys, 0);
    $merged['status_breakdown'] = [];
    $merged['top_products'] = [];
    $merged['platforms'] = [];

    $pcKeys = [
        'delivered_orders','delivered_qty','my_selling_total','meesho_selling_total',
        'gross_profit_margin','meesho_commission_cut','meesho_other_fee_cut',
        'meesho_total_earned','shipping_charged','tax_deducted','net_you_earned',
        'return_orders','return_qty','return_loss','return_shipping_cut',
        'rto_orders','rto_qty','rto_loss','rto_shipping_cut',
        'ads_cost','ads_gst','total_ads','final_in_hand','net_profit_loss','per_unit_profit',
    ];
    $merged['profit_calc'] = array_fill_keys($pcKeys, 0);

    $rbKeys = [
        'gross_sales_pieces','gross_sales_value','rto_pieces','rto_value',
        'customer_return_pieces','customer_return_value',
        'net_sold_pieces','net_product_revenue',
        'returned_stock_pieces','returned_stock_value',
        'rto_courier_charges','customer_return_shipping',
        'platform_fees_total','final_cash_in_hand',
        'cost_price_total','total_profit_loss'
    ];
    $merged['revenue_breakdown'] = array_fill_keys($rbKeys, 0);

    foreach ($results as $r) {
        foreach ($keys as $k) $merged[$k] += $r[$k] ?? 0;
        foreach ($pcKeys as $k) $merged['profit_calc'][$k] += ($r['profit_calc'][$k] ?? 0);
        foreach ($rbKeys as $k) $merged['revenue_breakdown'][$k] += ($r['revenue_breakdown'][$k] ?? 0);

        $p = $r['platform'] ?? 'Other';
        if (!isset($merged['platforms'][$p]))
            $merged['platforms'][$p] = ['files'=>0,'sold'=>0,'return'=>0,'rto'=>0,'payment'=>0,'profit'=>0];
        $merged['platforms'][$p]['files']++;
        $merged['platforms'][$p]['sold']    += $r['sold'] ?? 0;
        $merged['platforms'][$p]['return']  += $r['return'] ?? 0;
        $merged['platforms'][$p]['rto']     += $r['rto'] ?? 0;
        $merged['platforms'][$p]['payment'] += $r['payment_received'] ?? 0;
        $merged['platforms'][$p]['profit']  += $r['net_profit'] ?? 0;

        foreach ($r['status_breakdown'] ?? [] as $k => $v)
            $merged['status_breakdown'][$k] = ($merged['status_breakdown'][$k] ?? 0) + $v;
        foreach ($r['top_products'] ?? [] as $k => $v)
            $merged['top_products'][$k] = ($merged['top_products'][$k] ?? 0) + $v;
    }

    if ($merged['profit_calc']['delivered_qty'] > 0)
        $merged['profit_calc']['per_unit_profit'] = $merged['profit_calc']['net_profit_loss'] / $merged['profit_calc']['delivered_qty'];

    arsort($merged['top_products']);
    $merged['top_products'] = array_slice($merged['top_products'], 0, 10, true);
    return $merged;
}

// ========== UPLOAD ==========
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['upload_file'])) {
    $file = $_FILES['upload_file'];
    if ($file['error'] !== UPLOAD_ERR_OK) {
        $error = 'Upload failed (code ' . $file['error'] . ')';
    } else {
        $ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
        $allowed = ['xlsx','xls','csv','zip'];
        if (!in_array($ext, $allowed)) {
            $error = 'Invalid file type.';
        } else {
            $safeName = 'upload_' . date('Ymd_His') . '_' . uniqid() . '.' . $ext;
            $targetPath = $uploadDir . $safeName;
            if (move_uploaded_file($file['tmp_name'], $targetPath)) {
                try {
                    $filesToProcess = [];
                    if ($ext === 'zip') {
                        $zip = new ZipArchive();
                        if ($zip->open($targetPath) !== true) throw new Exception('Cannot open ZIP.');
                        $sessionExtract = $extractDir . 'zip_' . uniqid() . '/';
                        mkdir($sessionExtract, 0777, true);
                        $zip->extractTo($sessionExtract);
                        $zip->close();
                        $iterator = new RecursiveIteratorIterator(
                            new RecursiveDirectoryIterator($sessionExtract, RecursiveDirectoryIterator::SKIP_DOTS)
                        );
                        foreach ($iterator as $f) {
                            if (!$f->isFile()) continue;
                            if (strpos($f->getPathname(), '__MACOSX') !== false) continue;
                            if (strpos($f->getFilename(), '.') === 0) continue;
                            $fe = strtolower($f->getExtension());
                            if (in_array($fe, ['xlsx','xls','csv']))
                                $filesToProcess[] = ['path'=>$f->getPathname(),'name'=>$f->getFilename()];
                        }
                        if (empty($filesToProcess)) throw new Exception('No files in ZIP.');
                        $debug['extract_dir'] = $sessionExtract;
                    } else {
                        $filesToProcess[] = ['path'=>$targetPath,'name'=>$file['name']];
                    }

                    foreach ($filesToProcess as $fp) {
                        try {
                            $fileResults[] = analyzeFile($fp['path'], $fp['name'], $orderDetails, $selectedPlatform);
                        } catch (Exception $e) {
                            $fileResults[] = ['platform'=>'Other','filename'=>$fp['name'],'error'=>$e->getMessage()];
                        }
                    }
                    $successResults = array_filter($fileResults, fn($r) => !isset($r['error']));
                    if (empty($successResults)) {
                        $errs = array_map(fn($r) => $r['filename'] . ': ' . $r['error'], $fileResults);
                        throw new Exception("No files could be processed:\n" . implode("\n", $errs));
                    }
                    $summary = mergeResults($successResults);

                    if ($ext === 'zip' && !empty($debug['extract_dir'])) {
                        $it = new RecursiveIteratorIterator(
                            new RecursiveDirectoryIterator($debug['extract_dir'], RecursiveDirectoryIterator::SKIP_DOTS),
                            RecursiveIteratorIterator::CHILD_FIRST
                        );
                        foreach ($it as $f) $f->isDir() ? @rmdir($f->getPathname()) : @unlink($f->getPathname());
                        @rmdir($debug['extract_dir']);
                    }
                    @unlink($targetPath);
                } catch (Exception $e) { $error = 'Error: ' . $e->getMessage(); }
            } else { $error = 'Failed to save file.'; }
        }
    }
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Business Dashboard</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<style>
    /* ═══ THEME VARIABLES ═══ */
    :root {
        --bg-main: #f0f3f8;
        --bg-panel: #ffffff;
        --bg-panel-hover: #fafbfd;
        --bg-subtle: #f5f7fb;
        --border: #dbe2ec;
        --border-soft: #e8ecf3;
        --text-main: #3d4f60;
        --text-sub: #5a6b7c;
        --text-muted: #8895a7;
        --accent: #3b6fc4;
        --accent-hover: #2d5aa8;
        --topbar-1: #3b6fc4;
        --topbar-2: #2d5aa8;
        --table-head-bg: #f5f7fb;
        --shadow: rgba(0,0,0,0.06);
        --input-bg: #fafbfd;
    }
    html.dark {
        --bg-main: #0f1419;
        --bg-panel: #1a202c;
        --bg-panel-hover: #212936;
        --bg-subtle: #151b24;
        --border: #2d3748;
        --border-soft: #252d3a;
        --text-main: #e2e8f0;
        --text-sub: #a0aec0;
        --text-muted: #718096;
        --accent: #5b8bd8;
        --accent-hover: #4a78c2;
        --topbar-1: #1a202c;
        --topbar-2: #0f1419;
        --table-head-bg: #151b24;
        --shadow: rgba(0,0,0,0.4);
        --input-bg: #151b24;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }

    body {
        font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
        background: var(--bg-main);
        color: var(--text-main);
        font-size: 13px;
        min-height: 100vh;
        transition: background 0.3s, color 0.3s;
    }

    /* ═══ TOPBAR ═══ */
    .topbar {
        background: linear-gradient(135deg, var(--topbar-1) 0%, var(--topbar-2) 100%);
        color: #fff;
        height: 56px;
        display: flex;
        align-items: center;
        padding: 0 24px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        position: sticky;
        top: 0;
        z-index: 100;
    }
    .topbar .brand {
        display: flex;
        align-items: center;
        gap: 10px;
        font-weight: 700;
        font-size: 16px;
        padding-right: 20px;
        border-right: 1px solid rgba(255,255,255,0.2);
        margin-right: 20px;
    }
    .topbar .brand i { font-size: 18px; }
    .topbar .date-time {
        font-size: 12.5px;
        opacity: 0.9;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .topbar .nav-icons {
        margin-left: auto;
        display: flex;
        gap: 20px;
        align-items: center;
    }
    .topbar .nav-icons a {
        color: #fff;
        text-decoration: none;
        font-size: 12.5px;
        display: flex;
        align-items: center;
        gap: 6px;
        opacity: 0.9;
        transition: opacity 0.2s;
    }
    .topbar .nav-icons a:hover { opacity: 1; }
    .topbar .theme-toggle {
        background: rgba(255,255,255,0.15);
        border: none;
        color: #fff;
        width: 38px;
        height: 38px;
        border-radius: 50%;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 16px;
        transition: all 0.25s;
        margin-left: 10px;
    }
    .topbar .theme-toggle:hover {
        background: rgba(255,255,255,0.28);
        transform: rotate(20deg);
    }

    /* ═══ MAIN LAYOUT (no sidebar) ═══ */
    .main {
        max-width: 1500px;
        margin: 0 auto;
        padding: 24px 28px 40px;
    }

    /* ═══ TOP CALCULATION SUMMARY BAR ═══ */
    .calc-summary-bar {
        background: var(--bg-panel);
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 18px 22px;
        margin-bottom: 22px;
        box-shadow: 0 1px 3px var(--shadow);
    }
    .calc-summary-bar .csb-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 14px;
        padding-bottom: 12px;
        border-bottom: 1px solid var(--border-soft);
        flex-wrap: wrap;
        gap: 10px;
    }
    .calc-summary-bar .csb-title {
        font-size: 15px;
        font-weight: 700;
        color: var(--text-main);
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .calc-summary-bar .csb-title i { color: var(--accent); font-size: 18px; }
    .calc-summary-bar .csb-subtitle {
        font-size: 12px;
        color: var(--text-muted);
        font-weight: 500;
    }

    /* Medium-size calc grid */
    .calc-grid-medium {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 10px;
    }
    .calc-card-md {
        background: var(--bg-subtle);
        border: 1px solid var(--border-soft);
        border-radius: 6px;
        padding: 14px 14px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        min-height: 88px;
        transition: all 0.2s;
    }
    .calc-card-md:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px var(--shadow);
    }
    .calc-card-md .cc-label {
        font-size: 10.5px;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        font-weight: 700;
        color: var(--text-muted);
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 5px;
    }
    .calc-card-md .cc-label i { font-size: 11px; }
    .calc-card-md .cc-value {
        font-size: 19px;
        font-weight: 800;
        font-family: 'Courier New', monospace;
        letter-spacing: -0.5px;
        line-height: 1.1;
    }
    .calc-card-md .cc-desc {
        font-size: 10.5px;
        color: var(--text-muted);
        margin-top: 4px;
    }

    .calc-card-md.income { border-left: 4px solid #48bb78; }
    .calc-card-md.income .cc-value { color: #38a169; }

    .calc-card-md.expense { border-left: 4px solid #f56565; }
    .calc-card-md.expense .cc-value { color: #e53e3e; }

    .calc-card-md.highlight { border-left: 4px solid var(--accent); background: rgba(59,111,196,0.06); }
    .calc-card-md.highlight .cc-value { color: var(--accent); }

    .calc-card-md.gst { border-left: 4px solid #7e57c2; background: rgba(126,87,194,0.06); }
    .calc-card-md.gst .cc-value { color: #7e57c2; }

    .calc-card-md.tax { border-left: 4px solid #dd6b20; background: rgba(221,107,32,0.06); }
    .calc-card-md.tax .cc-value { color: #dd6b20; }

    .calc-card-md.final { border-left: 4px solid #3d4f60; background: rgba(61,79,96,0.08); }
    .calc-card-md.final.profit .cc-value { color: #38a169; }
    .calc-card-md.final.loss .cc-value { color: #e53e3e; }

    /* ═══ STAT CARDS ROW ═══ */
    .stat-cards {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 14px;
        margin-bottom: 22px;
    }
    .stat-card {
        border-radius: 8px;
        padding: 18px 20px;
        color: #fff;
        position: relative;
        overflow: hidden;
        min-height: 118px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    .stat-card .sc-label {
        font-size: 11.5px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        opacity: 0.95;
        margin-bottom: 4px;
    }
    .stat-card .sc-value {
        font-size: 26px;
        font-weight: 800;
        line-height: 1.1;
    }
    .stat-card .sc-sub {
        font-size: 11px;
        opacity: 0.9;
        margin-top: 6px;
    }
    .stat-card .sc-icon {
        position: absolute;
        right: 14px;
        top: 50%;
        transform: translateY(-50%);
        font-size: 64px;
        opacity: 0.15;
        pointer-events: none;
    }
    .stat-card.purple-1 { background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%); }
    .stat-card.purple-2 { background: linear-gradient(135deg, #a855f7 0%, #9333ea 100%); }
    .stat-card.purple-3 { background: linear-gradient(135deg, #7e57c2 0%, #673ab7 100%); }
    .stat-card.purple-4 { background: linear-gradient(135deg, #9575cd 0%, #7e57c2 100%); }
    .stat-card.green { background: linear-gradient(135deg, #48bb78 0%, #38a169 100%); }
    .stat-card.red { background: linear-gradient(135deg, #f56565 0%, #e53e3e 100%); }

    /* ═══ PANEL ═══ */
    .panel {
        background: var(--bg-panel);
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 22px 24px;
        margin-bottom: 20px;
        box-shadow: 0 1px 3px var(--shadow);
        transition: background 0.3s;
    }
    .panel-title {
        font-size: 15px;
        font-weight: 700;
        color: var(--text-main);
        padding-bottom: 12px;
        margin-bottom: 16px;
        border-bottom: 1px solid var(--border-soft);
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .panel-title i { color: var(--accent); }

    /* ═══ UPLOAD ZONE ═══ */
    .upload-zone {
        border: 2px dashed var(--border);
        border-radius: 8px;
        padding: 32px 24px;
        text-align: center;
        cursor: pointer;
        background: var(--input-bg);
        transition: all 0.25s;
    }
    .upload-zone:hover {
        border-color: var(--accent);
        background: rgba(59,111,196,0.06);
    }
    .upload-zone input[type="file"] { display: none; }
    .upload-zone .uz-icon {
        font-size: 38px;
        color: var(--accent);
        margin-bottom: 10px;
        display: block;
    }
    .upload-zone .uz-title {
        font-size: 15px;
        color: var(--text-main);
        font-weight: 700;
        margin-bottom: 4px;
    }
    .upload-zone .uz-sub {
        font-size: 12px;
        color: var(--text-muted);
    }

    /* Platform strip */
    .platform-strip {
        display: flex;
        gap: 8px;
        margin-bottom: 20px;
        flex-wrap: wrap;
        justify-content: center;
    }
    .platform-strip-btn {
        background: var(--bg-panel);
        border: 1px solid var(--border);
        border-radius: 6px;
        padding: 10px 18px;
        cursor: pointer;
        font-size: 12.5px;
        font-weight: 600;
        color: var(--text-sub);
        transition: all 0.2s;
        display: flex;
        align-items: center;
        gap: 6px;
        font-family: inherit;
    }
    .platform-strip-btn:hover {
        border-color: var(--accent);
        color: var(--accent);
    }
    .platform-strip-btn.active {
        background: linear-gradient(135deg, var(--accent), var(--accent-hover));
        color: #fff;
        border-color: var(--accent);
    }
    .platform-strip-btn i { font-size: 14px; }

    .btn-primary {
        background: linear-gradient(135deg, var(--accent), var(--accent-hover));
        color: #fff;
        border: none;
        padding: 12px 36px;
        border-radius: 6px;
        font-size: 14px;
        font-weight: 700;
        cursor: pointer;
        transition: all 0.2s;
        font-family: inherit;
        box-shadow: 0 2px 8px rgba(59,111,196,0.3);
        display: inline-flex;
        align-items: center;
        gap: 8px;
    }
    .btn-primary:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(59,111,196,0.4);
    }

    /* ═══ ERROR ═══ */
    .error-box {
        background: rgba(245,101,101,0.1);
        border-left: 4px solid #e53e3e;
        color: #c53030;
        padding: 14px 18px;
        border-radius: 6px;
        font-size: 13px;
        margin-bottom: 20px;
        font-family: monospace;
        white-space: pre-wrap;
    }

    /* ═══ TABLE ═══ */
    .table-wrap { overflow-x: auto; border-radius: 6px; border: 1px solid var(--border-soft); }
    .data-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
    .data-table thead th {
        background: var(--table-head-bg);
        color: var(--text-sub);
        padding: 12px 14px;
        text-align: left;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-weight: 700;
        border-bottom: 2px solid var(--border);
        white-space: nowrap;
    }
    .data-table tbody td {
        padding: 11px 14px;
        border-bottom: 1px solid var(--border-soft);
        color: var(--text-sub);
    }
    .data-table tbody td.num {
        text-align: right;
        font-family: 'Courier New', monospace;
        font-weight: 600;
    }
    .data-table tbody td.num.pos { color: #38a169; }
    .data-table tbody td.num.neg { color: #e53e3e; }
    .data-table tbody tr:hover { background: var(--bg-panel-hover); }
    .data-table tbody tr.delivered { background: rgba(72,187,120,0.05); }
    .data-table tbody tr.return { background: rgba(245,101,101,0.05); }
    .data-table tbody tr.rto { background: rgba(237,137,54,0.06); }
    .data-table tbody tr.shipped { background: rgba(66,153,225,0.05); }
    .data-table tbody tr.pending { background: rgba(159,122,234,0.05); }

    .status-badge {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 4px;
        font-size: 10.5px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.4px;
    }
    .status-badge.delivered { background: rgba(72,187,120,0.18); color: #38a169; }
    .status-badge.return { background: rgba(245,101,101,0.18); color: #e53e3e; }
    .status-badge.rto { background: rgba(237,137,54,0.18); color: #dd6b20; }
    .status-badge.shipped { background: rgba(66,153,225,0.18); color: #3182ce; }
    .status-badge.cancelled { background: rgba(160,174,192,0.18); color: #718096; }
    .status-badge.pending { background: rgba(159,122,234,0.18); color: #805ad5; }
    .status-badge.other, .status-badge.unknown { background: rgba(160,174,192,0.12); color: #718096; }

    /* ═══ FLOW ROW ═══ */
    .flow-row {
        display: flex;
        align-items: center;
        gap: 14px;
        padding: 14px 18px;
        border-radius: 6px;
        margin-bottom: 8px;
        background: var(--bg-subtle);
        border-left: 4px solid var(--border);
    }
    .flow-row.income { border-left-color: #48bb78; background: rgba(72,187,120,0.05); }
    .flow-row.expense { border-left-color: #f56565; background: rgba(245,101,101,0.05); }
    .flow-row.highlight { border-left-color: var(--accent); background: rgba(59,111,196,0.06); }
    .flow-row.gst { border-left-color: #7e57c2; background: rgba(126,87,194,0.05); }
    .flow-row.itc { border-left-color: #48bb78; background: rgba(72,187,120,0.06); }
    .flow-row.tax { border-left-color: #dd6b20; background: rgba(221,107,32,0.05); }
    .flow-row.final {
        background: linear-gradient(135deg, #3d4f60, #2d3d4f);
        color: #fff;
        border-left-color: var(--accent);
        padding: 18px 20px;
        margin-top: 14px;
    }
    .flow-row.final.profit { border-left-color: #48bb78; }
    .flow-row.final.loss { border-left-color: #f56565; }

    .flow-row .fr-badge {
        width: 38px;
        height: 38px;
        border-radius: 6px;
        background: var(--bg-panel);
        color: var(--text-sub);
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 14px;
        font-weight: 800;
        flex-shrink: 0;
    }
    .flow-row.income .fr-badge { background: rgba(72,187,120,0.2); color: #38a169; }
    .flow-row.expense .fr-badge { background: rgba(245,101,101,0.2); color: #e53e3e; }
    .flow-row.highlight .fr-badge { background: rgba(59,111,196,0.2); color: var(--accent); }
    .flow-row.gst .fr-badge { background: rgba(126,87,194,0.2); color: #7e57c2; }
    .flow-row.itc .fr-badge { background: rgba(72,187,120,0.2); color: #38a169; }
    .flow-row.tax .fr-badge { background: rgba(221,107,32,0.2); color: #dd6b20; }
    .flow-row.final .fr-badge { background: rgba(255,255,255,0.15); color: #fff; width: 46px; height: 46px; font-size: 20px; }

    .flow-row .fr-content { flex: 1; min-width: 0; }
    .flow-row .fr-title { font-size: 13.5px; font-weight: 700; color: var(--text-main); margin-bottom: 3px; }
    .flow-row.final .fr-title { color: #fff; font-size: 15px; }
    .flow-row .fr-desc { font-size: 11.5px; color: var(--text-muted); }
    .flow-row.final .fr-desc { color: rgba(255,255,255,0.75); }
    .flow-row .fr-amount {
        font-family: 'Courier New', monospace;
        font-weight: 800;
        font-size: 16px;
        color: var(--text-main);
        text-align: right;
        flex-shrink: 0;
    }
    .flow-row.income .fr-amount { color: #38a169; }
    .flow-row.expense .fr-amount { color: #e53e3e; }
    .flow-row.highlight .fr-amount { color: var(--accent); }
    .flow-row.gst .fr-amount { color: #7e57c2; }
    .flow-row.itc .fr-amount { color: #38a169; }
    .flow-row.tax .fr-amount { color: #dd6b20; }
    .flow-row.final .fr-amount { color: #fff; font-size: 22px; }
    .flow-row.final.profit .fr-amount { color: #c6f6d5; }
    .flow-row.final.loss .fr-amount { color: #fed7d7; }

    /* ═══ INFO NOTE ═══ */
    .info-note {
        background: rgba(59,111,196,0.06);
        border-left: 4px solid var(--accent);
        padding: 12px 16px;
        border-radius: 6px;
        font-size: 12.5px;
        color: var(--text-sub);
        margin-bottom: 16px;
        line-height: 1.6;
    }
    .info-note.purple { background: rgba(126,87,194,0.06); border-left-color: #7e57c2; }

    .footer {
        text-align: center;
        color: var(--text-muted);
        font-size: 12px;
        padding: 24px;
    }

    /* ═══ RESPONSIVE ═══ */
    @media (max-width: 1200px) {
        .calc-grid-medium { grid-template-columns: repeat(4, 1fr); }
    }
    @media (max-width: 900px) {
        .stat-cards { grid-template-columns: repeat(2, 1fr); }
        .calc-grid-medium { grid-template-columns: repeat(3, 1fr); }
        .main { padding: 16px; }
    }
    @media (max-width: 600px) {
        .topbar { padding: 0 14px; }
        .topbar .date-time { display: none; }
        .topbar .nav-icons a span { display: none; }
        .stat-cards { grid-template-columns: 1fr; }
        .calc-grid-medium { grid-template-columns: repeat(2, 1fr); }
        .calc-card-md .cc-value { font-size: 16px; }
    }
</style>
</head>
<body>

<!-- ═══ TOPBAR with DARK MODE TOGGLE ═══ -->
<div class="topbar">
    <div class="brand">
        <i class="fas fa-chart-line"></i> Business Report
    </div>
    <div class="date-time">
        <i class="fas fa-calendar-alt"></i>
        <?= date('D d M Y h:i:s A') ?>
    </div>
    <div class="nav-icons">
        <a href="#uploadCard"><i class="fas fa-cloud-upload-alt"></i> <span>Upload</span></a>
        <a href="#calcSummary"><i class="fas fa-calculator"></i> <span>Calculate</span></a>
        <a href="#ordersTable"><i class="fas fa-list"></i> <span>Orders</span></a>
        <button class="theme-toggle" onclick="toggleDarkMode()" title="Toggle Dark Mode" id="themeToggle">
            <i class="fas fa-moon"></i>
        </button>
    </div>
</div>

<!-- ═══ MAIN ═══ -->
<div class="main">

    <?php if ($error): ?>
        <div class="error-box">⚠️ <?= htmlspecialchars($error) ?></div>
    <?php endif; ?>

    <!-- ═══ UPLOAD PANEL ═══ -->
    <div class="panel" id="uploadCard">
        <form method="POST" enctype="multipart/form-data" id="uploadForm">
            <div class="platform-strip">
                <button type="button" class="platform-strip-btn <?= $selectedPlatform === 'meesho' ? 'active' : '' ?>" data-platform="meesho" onclick="selectPlatform('meesho')">
                    <i class="fas fa-shopping-bag"></i> Meesho
                </button>
                <button type="button" class="platform-strip-btn <?= $selectedPlatform === 'flipkart' ? 'active' : '' ?>" data-platform="flipkart" onclick="selectPlatform('flipkart')">
                    <i class="fas fa-shopping-cart"></i> Flipkart
                </button>
                <button type="button" class="platform-strip-btn <?= $selectedPlatform === 'amazon' ? 'active' : '' ?>" data-platform="amazon" onclick="selectPlatform('amazon')">
                    <i class="fas fa-box"></i> Amazon
                </button>
                <button type="button" class="platform-strip-btn <?= $selectedPlatform === 'myntra' ? 'active' : '' ?>" data-platform="myntra" onclick="selectPlatform('myntra')">
                    <i class="fas fa-tshirt"></i> Myntra
                </button>
                <button type="button" class="platform-strip-btn <?= $selectedPlatform === 'auto' ? 'active' : '' ?>" data-platform="auto" onclick="selectPlatform('auto')">
                    <i class="fas fa-magic"></i> Auto Detect
                </button>
            </div>

            <input type="hidden" name="platform" id="platformInput" value="<?= htmlspecialchars($selectedPlatform) ?>">

            <div class="upload-zone" onclick="document.getElementById('fileInput').click()">
                <span class="uz-icon"><i class="fas fa-cloud-upload-alt"></i></span>
                <div class="uz-title" id="uploadTitle">
                    <?php
                    $titles = [
                        'meesho'   => 'Upload your Meesho payment file',
                        'flipkart' => 'Upload your Flipkart payment file',
                        'amazon'   => 'Upload your Amazon payment file',
                        'myntra'   => 'Upload your Myntra payment file',
                        'auto'     => 'Upload your payment file (auto-detect)',
                    ];
                    echo $titles[$selectedPlatform] ?? 'Upload your payment file';
                    ?>
                </div>
                <div class="uz-sub">Supported: .xlsx / .xls / .csv / .zip</div>
                <input type="file" name="upload_file" id="fileInput" accept=".xlsx,.xls,.csv,.zip" required onclick="event.stopPropagation()">
            </div>

            <div style="text-align: center; margin-top: 20px;">
                <button type="submit" class="btn-primary"><i class="fas fa-chart-bar"></i> Generate Report</button>
            </div>
        </form>
    </div>

    <?php if ($summary !== null):
        $pc = $summary['profit_calc'];
        $rb = $summary['revenue_breakdown'];

        $platformLabel = 'Business';
        if (!empty($fileResults)) {
            foreach ($fileResults as $fr) {
                if (!isset($fr['error']) && !empty($fr['platform'])) {
                    $platformLabel = $fr['platform'];
                    break;
                }
            }
        }

        $grossSales       = $rb['gross_sales_value'];
        $rtoValue         = abs($rb['rto_value']);
        $customerReturn   = abs($rb['customer_return_value']);
        $rtoCourier       = abs($rb['rto_courier_charges']);
        $customerRetShip  = abs($rb['customer_return_shipping']);
        $returnCharges    = $rtoCourier + $customerRetShip;
        $adsCost          = abs($summary['ads_total']);
        $netSales         = $rb['net_product_revenue'];
        $platformFees     = $rb['platform_fees_total'];
        $meeshoCommission = abs($summary['meesho_commission']);
        $costPrice        = $rb['cost_price_total'];

        $finalCash        = $netSales - $adsCost - $returnCharges;
        $finalProfit      = $finalCash - $platformFees - $costPrice;
        $isProfit = $finalProfit >= 0;

        $gst = calculateIndianGST($summary);

        $profitBeforeTax = $finalProfit;
        $totalTaxLiability = max(0, $gst['net_gst_payable']) + $gst['tcs'] + $gst['tds'];
        $profitAfterTax = $profitBeforeTax - $totalTaxLiability;
        $isAfterTaxProfit = $profitAfterTax >= 0;
    ?>

    <!-- ═══ FULL CALCULATION SUMMARY (Top Medium Boxes) ═══ -->
    <div class="calc-summary-bar" id="calcSummary">
        <div class="csb-header">
            <div class="csb-title">
                <i class="fas fa-calculator"></i> Full Calculation · <?= htmlspecialchars($platformLabel) ?>
            </div>
            <div class="csb-subtitle">
                <i class="fas fa-info-circle"></i> All amounts in ₹ · GST rate 12%
            </div>
        </div>

        <div class="calc-grid-medium">
            <div class="calc-card-md income">
                <div class="cc-label"><i class="fas fa-shopping-cart"></i> Gross Sales</div>
                <div class="cc-value">₹<?= number_format($grossSales, 0) ?></div>
                <div class="cc-desc"><?= number_format($rb['gross_sales_pieces']) ?> pieces</div>
            </div>

            <div class="calc-card-md expense">
                <div class="cc-label"><i class="fas fa-undo"></i> Less: RTO</div>
                <div class="cc-value">₹<?= number_format($rtoValue, 0) ?></div>
                <div class="cc-desc"><?= number_format($rb['rto_pieces']) ?> pieces</div>
            </div>

            <div class="calc-card-md expense">
                <div class="cc-label"><i class="fas fa-reply"></i> Less: Return</div>
                <div class="cc-value">₹<?= number_format($customerReturn, 0) ?></div>
                <div class="cc-desc"><?= number_format($rb['customer_return_pieces']) ?> pieces</div>
            </div>

            <div class="calc-card-md highlight">
                <div class="cc-label"><i class="fas fa-equals"></i> Net Sales</div>
                <div class="cc-value">₹<?= number_format($netSales, 0) ?></div>
                <div class="cc-desc"><?= number_format($rb['net_sold_pieces']) ?> kept</div>
            </div>

            <div class="calc-card-md expense">
                <div class="cc-label"><i class="fas fa-truck"></i> Less: Ret. Ship</div>
                <div class="cc-value">₹<?= number_format($returnCharges, 0) ?></div>
                <div class="cc-desc">Courier charges</div>
            </div>

            <div class="calc-card-md expense">
                <div class="cc-label"><i class="fas fa-bullhorn"></i> Less: Ads</div>
                <div class="cc-value">₹<?= number_format($adsCost, 0) ?></div>
                <div class="cc-desc"><?= $summary['ads_rows'] ?> entries</div>
            </div>

            <div class="calc-card-md highlight">
                <div class="cc-label"><i class="fas fa-wallet"></i> Cash in Hand</div>
                <div class="cc-value">₹<?= number_format($finalCash, 0) ?></div>
                <div class="cc-desc">After ads + returns</div>
            </div>

            <div class="calc-card-md expense">
                <div class="cc-label"><i class="fas fa-percent"></i> Less: Fees</div>
                <div class="cc-value">₹<?= number_format($platformFees, 0) ?></div>
                <div class="cc-desc">Platform charges</div>
            </div>

            <div class="calc-card-md gst">
                <div class="cc-label"><i class="fas fa-file-invoice-dollar"></i> GST Payable</div>
                <div class="cc-value">₹<?= number_format(abs($gst['net_gst_payable']), 0) ?></div>
                <div class="cc-desc">After ITC</div>
            </div>

            <div class="calc-card-md final <?= $isProfit ? 'profit' : 'loss' ?>">
                <div class="cc-label"><i class="fas fa-trophy"></i> Final <?= $isProfit ? 'Profit' : 'Loss' ?></div>
                <div class="cc-value">₹<?= number_format(abs($finalProfit), 0) ?></div>
                <div class="cc-desc">Before tax</div>
            </div>
        </div>
    </div>

    <!-- ═══ STAT CARDS (Purple gradients) ═══ -->
    <div class="stat-cards">
        <div class="stat-card <?= $isAfterTaxProfit ? 'purple-1' : 'red' ?>">
            <div class="sc-label">Profit After Tax</div>
            <div class="sc-value">₹<?= number_format(abs($profitAfterTax), 0) ?></div>
            <div class="sc-sub">Before Tax: ₹<?= number_format(abs($profitBeforeTax), 0) ?></div>
            <i class="fas fa-chart-line sc-icon"></i>
        </div>

        <div class="stat-card purple-2">
            <div class="sc-label">Total Orders</div>
            <div class="sc-value"><?= number_format($summary['total_orders']) ?></div>
            <div class="sc-sub"><?= number_format($summary['total_rows']) ?> total rows</div>
            <i class="fas fa-shopping-bag sc-icon"></i>
        </div>

        <div class="stat-card purple-3">
            <div class="sc-label">GST Payable (Net)</div>
            <div class="sc-value">₹<?= number_format(abs($gst['net_gst_payable']), 0) ?></div>
            <div class="sc-sub">Output: ₹<?= number_format($gst['output_gst'], 0) ?></div>
            <i class="fas fa-file-invoice-dollar sc-icon"></i>
        </div>

        <div class="stat-card purple-4">
            <div class="sc-label">Total Tax</div>
            <div class="sc-value">₹<?= number_format($totalTaxLiability, 0) ?></div>
            <div class="sc-sub">GST + TCS + TDS</div>
            <i class="fas fa-receipt sc-icon"></i>
        </div>
    </div>

    <!-- ═══ REVENUE BREAKDOWN ═══ -->
    <div class="panel" id="revenueSection">
        <div class="panel-title"><i class="fas fa-chart-pie"></i> Sales & Revenue Breakdown</div>

        <div class="flow-row income">
            <div class="fr-badge">1</div>
            <div class="fr-content">
                <div class="fr-title">Gross Sales</div>
                <div class="fr-desc"><?= number_format($rb['gross_sales_pieces']) ?> pieces ordered</div>
            </div>
            <div class="fr-amount">₹<?= number_format($grossSales, 2) ?></div>
        </div>

        <div class="flow-row expense">
            <div class="fr-badge">2</div>
            <div class="fr-content">
                <div class="fr-title">Less: RTO Return Amount</div>
                <div class="fr-desc"><?= number_format($rb['rto_pieces']) ?> pieces returned to origin</div>
            </div>
            <div class="fr-amount">− ₹<?= number_format($rtoValue, 2) ?></div>
        </div>

        <div class="flow-row expense">
            <div class="fr-badge">3</div>
            <div class="fr-content">
                <div class="fr-title">Less: Customer Return</div>
                <div class="fr-desc"><?= number_format($rb['customer_return_pieces']) ?> pieces returned by customers</div>
            </div>
            <div class="fr-amount">− ₹<?= number_format($customerReturn, 2) ?></div>
        </div>

        <div class="flow-row highlight">
            <div class="fr-badge">=</div>
            <div class="fr-content">
                <div class="fr-title">Net Product Revenue (Main Sell)</div>
                <div class="fr-desc"><?= number_format($rb['net_sold_pieces']) ?> pieces kept by customers</div>
            </div>
            <div class="fr-amount">₹<?= number_format($netSales, 2) ?></div>
        </div>

        <div class="flow-row expense">
            <div class="fr-badge">4</div>
            <div class="fr-content">
                <div class="fr-title">Less: Return Charges (Courier)</div>
                <div class="fr-desc">RTO ₹<?= number_format($rtoCourier, 2) ?> + Return ₹<?= number_format($customerRetShip, 2) ?></div>
            </div>
            <div class="fr-amount">− ₹<?= number_format($returnCharges, 2) ?></div>
        </div>

        <div class="flow-row expense">
            <div class="fr-badge">5</div>
            <div class="fr-content">
                <div class="fr-title">Less: Ads Cost</div>
                <div class="fr-desc"><?= $summary['ads_rows'] ?> entries</div>
            </div>
            <div class="fr-amount">− ₹<?= number_format($adsCost, 2) ?></div>
        </div>

        <div class="flow-row highlight">
            <div class="fr-badge">=</div>
            <div class="fr-content">
                <div class="fr-title">Final Cash in Hand</div>
                <div class="fr-desc">Net Sales − Return Charges − Ads Cost</div>
            </div>
            <div class="fr-amount">₹<?= number_format($finalCash, 2) ?></div>
        </div>

        <div class="flow-row expense">
            <div class="fr-badge">6</div>
            <div class="fr-content">
                <div class="fr-title">Less: Platform Fees & Commission</div>
                <div class="fr-desc">All marketplace charges</div>
            </div>
            <div class="fr-amount">− ₹<?= number_format($platformFees, 2) ?></div>
        </div>

        <div class="flow-row final <?= $isProfit ? 'profit' : 'loss' ?>">
            <div class="fr-badge">🏁</div>
            <div class="fr-content">
                <div class="fr-title">Total <?= $isProfit ? 'Profit' : 'Loss' ?></div>
                <div class="fr-desc"><?= $isProfit ? 'Profit — Great job!' : 'Loss — Review pricing' ?></div>
            </div>
            <div class="fr-amount">₹<?= number_format(abs($finalProfit), 2) ?></div>
        </div>
    </div>

    <!-- ═══ GST FILING ═══ -->
    <div class="panel" id="gstSection">
        <div class="panel-title"><i class="fas fa-file-invoice-dollar"></i> GST Filing Summary</div>

        <div class="info-note purple">
            <i class="fas fa-info-circle"></i> Output GST − Input Tax Credit (ITC) = Net GST Payable · Rate: 12%
        </div>

        <div class="flow-row gst">
            <div class="fr-badge">OUT</div>
            <div class="fr-content">
                <div class="fr-title">Output GST (GST on Sales)</div>
                <div class="fr-desc">Turnover ₹<?= number_format($gst['taxable_turnover'], 2) ?> × 12%</div>
            </div>
            <div class="fr-amount">₹<?= number_format($gst['output_gst'], 2) ?></div>
        </div>

        <div class="flow-row gst">
            <div class="fr-badge">CG</div>
            <div class="fr-content"><div class="fr-title">Central GST (CGST)</div><div class="fr-desc">Intra-state</div></div>
            <div class="fr-amount">₹<?= number_format($gst['output_cgst'], 2) ?></div>
        </div>

        <div class="flow-row gst">
            <div class="fr-badge">SG</div>
            <div class="fr-content"><div class="fr-title">State GST (SGST)</div><div class="fr-desc">Intra-state</div></div>
            <div class="fr-amount">₹<?= number_format($gst['output_sgst'], 2) ?></div>
        </div>

        <div class="flow-row gst">
            <div class="fr-badge">IG</div>
            <div class="fr-content"><div class="fr-title">Integrated GST (IGST)</div><div class="fr-desc">Inter-state</div></div>
            <div class="fr-amount">₹<?= number_format($gst['output_igst'], 2) ?></div>
        </div>

        <div class="flow-row itc">
            <div class="fr-badge">ITC</div>
            <div class="fr-content">
                <div class="fr-title">Input Tax Credit (Total)</div>
                <div class="fr-desc">GST already paid on expenses</div>
            </div>
            <div class="fr-amount">₹<?= number_format($gst['total_itc'], 2) ?></div>
        </div>

        <?php
        $itcItems = [
            'itc_on_ads' => ['ADS', 'ITC on Ads'],
            'itc_on_commission' => ['COM', 'ITC on Commission'],
            'itc_on_shipping' => ['SHP', 'ITC on Shipping'],
            'itc_on_warehouse' => ['WHS', 'ITC on Warehousing'],
            'itc_on_fixed_fee' => ['FIX', 'ITC on Fixed Fees'],
            'itc_on_other' => ['OTH', 'ITC on Other Charges'],
        ];
        foreach ($itcItems as $key => $info):
            if ($gst[$key] > 0): ?>
                <div class="flow-row itc" style="margin-left:24px;">
                    <div class="fr-badge"><?= $info[0] ?></div>
                    <div class="fr-content"><div class="fr-title"><?= $info[1] ?></div></div>
                    <div class="fr-amount">₹<?= number_format($gst[$key], 2) ?></div>
                </div>
        <?php endif; endforeach; ?>

        <div class="flow-row final <?= $gst['net_gst_payable'] >= 0 ? 'loss' : 'profit' ?>">
            <div class="fr-badge">💳</div>
            <div class="fr-content">
                <div class="fr-title"><?= $gst['net_gst_payable'] >= 0 ? 'Net GST Payable' : 'Net GST Credit' ?></div>
                <div class="fr-desc">Output ₹<?= number_format($gst['output_gst'], 2) ?> − ITC ₹<?= number_format($gst['total_itc'], 2) ?></div>
            </div>
            <div class="fr-amount">₹<?= number_format(abs($gst['net_gst_payable']), 2) ?></div>
        </div>
    </div>

    <!-- ═══ TAX SUMMARY ═══ -->
    <div class="panel" id="taxSection">
        <div class="panel-title"><i class="fas fa-receipt"></i> Total Tax Summary</div>

        <div class="flow-row tax">
            <div class="fr-badge">GST</div>
            <div class="fr-content"><div class="fr-title">Net GST Payable (after ITC)</div></div>
            <div class="fr-amount">₹<?= number_format(abs($gst['net_gst_payable']), 2) ?></div>
        </div>

        <?php if ($gst['tcs'] > 0): ?>
        <div class="flow-row tax">
            <div class="fr-badge">TCS</div>
            <div class="fr-content"><div class="fr-title">TCS (Tax Collected at Source)</div></div>
            <div class="fr-amount">₹<?= number_format($gst['tcs'], 2) ?></div>
        </div>
        <?php endif; ?>

        <?php if ($gst['tds'] > 0): ?>
        <div class="flow-row tax">
            <div class="fr-badge">TDS</div>
            <div class="fr-content"><div class="fr-title">TDS (Tax Deducted at Source)</div></div>
            <div class="fr-amount">₹<?= number_format($gst['tds'], 2) ?></div>
        </div>
        <?php endif; ?>

        <div class="flow-row final loss">
            <div class="fr-badge">🧾</div>
            <div class="fr-content"><div class="fr-title">Total Tax Liability</div><div class="fr-desc">Payable to government</div></div>
            <div class="fr-amount">₹<?= number_format($totalTaxLiability, 2) ?></div>
        </div>

        <div class="flow-row final <?= $isAfterTaxProfit ? 'profit' : 'loss' ?>" style="margin-top:16px;">
            <div class="fr-badge">💎</div>
            <div class="fr-content">
                <div class="fr-title">Profit After Tax (Take Home)</div>
                <div class="fr-desc">Before Tax ₹<?= number_format($profitBeforeTax, 2) ?> − Total Tax ₹<?= number_format($totalTaxLiability, 2) ?></div>
            </div>
            <div class="fr-amount">₹<?= number_format(abs($profitAfterTax), 2) ?></div>
        </div>
    </div>

    <!-- ═══ ORDERS TABLE ═══ -->
    <?php if (!empty($orderDetails)): ?>
    <div class="panel" id="ordersTable">
        <div class="panel-title"><i class="fas fa-list-alt"></i> Every Product — Line by Line (<?= count($orderDetails) ?> orders)</div>
        <div class="table-wrap">
            <table class="data-table" style="min-width: 1900px;">
                <thead>
                    <tr>
                        <th>#</th><th>Order No</th><th>Product</th><th>Qty</th><th>Status</th>
                        <th>My Price</th><th>Mkt Price</th><th>My Total</th><th>Mkt Total</th>
                        <th>Margin</th><th>Commission</th><th>Gold</th><th>Mall</th><th>Fixed</th>
                        <th>Warehouse</th><th>Shipping</th><th>Return Ship</th><th>Return Prem</th>
                        <th>Support</th><th>TCS</th><th>TDS</th><th>Final Settlement</th>
                    </tr>
                </thead>
                <tbody>
                    <?php $idx = 1; foreach ($orderDetails as $o): ?>
                        <tr class="<?= $o['status_type'] ?>">
                            <td><?= $idx++ ?></td>
                            <td style="font-family:monospace;font-size:11px;"><?= htmlspecialchars($o['order_no']) ?></td>
                            <td>
                                <strong style="color:var(--text-main);"><?= htmlspecialchars(mb_strimwidth($o['product'], 0, 40, '...')) ?></strong>
                                <?php if (!empty($o['sku'])): ?>
                                    <div style="font-size:10px;color:var(--text-muted);margin-top:2px;font-family:monospace;">SKU: <?= htmlspecialchars($o['sku']) ?></div>
                                <?php endif; ?>
                            </td>
                            <td class="num"><?= $o['qty'] ?></td>
                            <td><span class="status-badge <?= $o['status_type'] ?>"><?= htmlspecialchars($o['status_raw'] ?: 'Unknown') ?></span></td>
                            <td class="num">₹<?= number_format($o['my_price'], 2) ?></td>
                            <td class="num">₹<?= number_format($o['meesho_price'], 2) ?></td>
                            <td class="num">₹<?= number_format($o['my_total'], 2) ?></td>
                            <td class="num">₹<?= number_format($o['meesho_total'], 2) ?></td>
                            <td class="num <?= $o['margin'] >= 0 ? 'pos' : 'neg' ?>">₹<?= number_format($o['margin'], 2) ?></td>
                            <td class="num <?= $o['commission'] < 0 ? 'neg' : '' ?>">₹<?= number_format($o['commission'], 2) ?></td>
                            <td class="num <?= $o['gold_fee'] < 0 ? 'neg' : '' ?>">₹<?= number_format($o['gold_fee'], 2) ?></td>
                            <td class="num <?= $o['mall_fee'] < 0 ? 'neg' : '' ?>">₹<?= number_format($o['mall_fee'], 2) ?></td>
                            <td class="num <?= $o['fixed_fee'] < 0 ? 'neg' : '' ?>">₹<?= number_format($o['fixed_fee'], 2) ?></td>
                            <td class="num <?= $o['warehousing_fee'] < 0 ? 'neg' : '' ?>">₹<?= number_format($o['warehousing_fee'], 2) ?></td>
                            <td class="num <?= $o['shipping_charge'] < 0 ? 'neg' : '' ?>">₹<?= number_format($o['shipping_charge'], 2) ?></td>
                            <td class="num <?= $o['return_shipping'] < 0 ? 'neg' : '' ?>">₹<?= number_format($o['return_shipping'], 2) ?></td>
                            <td class="num <?= $o['return_premium'] < 0 ? 'neg' : '' ?>">₹<?= number_format($o['return_premium'], 2) ?></td>
                            <td class="num <?= $o['other_support'] < 0 ? 'neg' : '' ?>">₹<?= number_format($o['other_support'], 2) ?></td>
                            <td class="num <?= $o['tcs'] < 0 ? 'neg' : '' ?>">₹<?= number_format($o['tcs'], 2) ?></td>
                            <td class="num <?= $o['tds'] < 0 ? 'neg' : '' ?>">₹<?= number_format($o['tds'], 2) ?></td>
                            <td class="num <?= $o['final_settlement'] >= 0 ? 'pos' : 'neg' ?>" style="font-weight:800;">₹<?= number_format($o['final_settlement'], 2) ?></td>
                        </tr>
                    <?php endforeach; ?>
                </tbody>
            </table>
        </div>
    </div>
    <?php endif; ?>

    <?php endif; ?>

    <div class="footer">
        Business Report · Meesho · Flipkart · Amazon · Myntra · GST Filing
    </div>
</div>

<script>
// ═══ DARK MODE TOGGLE ═══
function toggleDarkMode() {
    const html = document.documentElement;
    const btn = document.getElementById('themeToggle');
    const icon = btn.querySelector('i');

    if (html.classList.contains('dark')) {
        html.classList.remove('dark');
        icon.className = 'fas fa-moon';
        localStorage.setItem('theme', 'light');
    } else {
        html.classList.add('dark');
        icon.className = 'fas fa-sun';
        localStorage.setItem('theme', 'dark');
    }
}

// Load saved theme
(function() {
    const saved = localStorage.getItem('theme');
    if (saved === 'dark') {
        document.documentElement.classList.add('dark');
        const icon = document.querySelector('#themeToggle i');
        if (icon) icon.className = 'fas fa-sun';
    }
})();

// ═══ Platform selection ═══
function selectPlatform(platform) {
    document.getElementById('platformInput').value = platform;
    document.querySelectorAll('.platform-strip-btn').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.platform === platform);
    });
    const titles = {
        'meesho':   'Upload your Meesho payment file',
        'flipkart': 'Upload your Flipkart payment file',
        'amazon':   'Upload your Amazon payment file',
        'myntra':   'Upload your Myntra payment file',
        'auto':     'Upload your payment file (auto-detect)'
    };
    const el = document.getElementById('uploadTitle');
    if (el) el.textContent = titles[platform] || 'Upload your payment file';
}

// ═══ Drag & drop ═══
const ua = document.querySelector('.upload-zone');
const fi = document.getElementById('fileInput');
if (ua && fi) {
    ['dragenter','dragover'].forEach(evt => ua.addEventListener(evt, e => {
        e.preventDefault();
        ua.style.borderColor = 'var(--accent)';
        ua.style.background = 'rgba(59,111,196,0.08)';
    }));
    ['dragleave','drop'].forEach(evt => ua.addEventListener(evt, e => {
        e.preventDefault();
        ua.style.borderColor = 'var(--border)';
        ua.style.background = 'var(--input-bg)';
    }));
    ua.addEventListener('drop', e => {
        if (e.dataTransfer.files.length) fi.files = e.dataTransfer.files;
    });
}
</script>
</body>
</html>