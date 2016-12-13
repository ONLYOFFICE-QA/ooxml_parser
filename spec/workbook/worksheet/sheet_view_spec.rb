require 'spec_helper'

describe OoxmlParser::Pane do
  it 'freeze_panes.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/sheet_view/freeze_panes.xlsx')
    expect(xlsx.worksheets.first.sheet_views.first.pane.top_left_cell).to eq(OoxmlParser::Coordinates.new(3, 'B'))
  end
end
