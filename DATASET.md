# Dataset Documentation

## Overview

The Sales Dashboard is built on the **Sample - Superstore** dataset, a comprehensive retail transactions dataset containing 9,995+ sales records spanning 2014-2017.

## Data Structure

### Orders Table (9,995 records)
Core transaction data for all sales:
- **Row ID**: Unique identifier for each record
- **Order ID**: Unique order identifier
- **Order Date**: Date of purchase (2014-2017)
- **Ship Date**: Date of shipment
- **Ship Mode**: Shipping method (Standard, First Class, Second Class, Same Day)
- **Customer ID**: Unique customer identifier
- **Customer Name**: Customer name
- **Segment**: Customer segment (Consumer, Corporate, Home Office)
- **Geography**: Country, City, State, Postal Code, Region
- **Product**: Product ID, Category, Sub-Category, Product Name
- **Metrics**: Sales, Quantity, Discount, Profit

### People Table (5 records)
Manager-to-region mapping for accountability tracking:
- **Person**: Manager name
- **Region**: Assigned region (Central, East, South, West, West)

### Returns Table (297 records)
Product returns linked to original orders:
- **Returned**: Return status indicator
- **Order ID**: Links to original order for tracking return patterns

## Key Metrics

| Metric | Type | Description |
|--------|------|-------------|
| Sales | Continuous | Total revenue from transaction |
| Profit | Continuous | Net profit (Sales - Costs - Discount impact) |
| Discount | Continuous | Discount percentage applied to order |
| Quantity | Discrete | Number of items ordered |
| Return Rate | Derived | Percentage of orders returned by region/category |

## Dimensions

**Geographic**: Region (Central, East, South, West), State, City, Postal Code

**Temporal**: Order Date, Ship Date (2014-2017)

**Business**: Segment (Consumer, Corporate, Home Office), Category (Furniture, Office Supplies, Technology), Sub-Category (11 types)

**Hierarchical Drill-Down**: Region → Segment → Category → Sub-Category → Product

## Data Quality

- **Completeness**: 100% - no missing values
- **Accuracy**: Validated order-to-return relationships
- **Time Coverage**: 4 years (2014-2017)
- **Granularity**: Transaction-level detail
- **Geographic Coverage**: United States

## File Source

- **Source**: Tableau Sample Dataset (Superstore)
- **Format**: Excel (.xls)
- **Records**: 9,995 orders + 5 people + 297 returns
- **Size**: ~2 MB

## Usage

This dataset enables:
- Regional and segment performance analysis
- Profitability trend identification
- Discount impact assessment
- Return pattern analysis for quality tracking
- Year-over-year growth comparison
- Manager accountability through region-level metrics

## Notes

- Return data spans multiple years but focuses on quality indicators
- Discount varies by order and affects profitability
- Seasonal patterns visible in order dates
- All data is historical and demonstrative
