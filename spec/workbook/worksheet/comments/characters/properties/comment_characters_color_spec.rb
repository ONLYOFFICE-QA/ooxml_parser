require 'spec_helper'

describe 'comment_characters_color' do
  it 'color_unknown.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/comments/characters/properties/color/color_unknown.xlsx')
    expect(xlsx.worksheets.first.comments.comments.first.characters.first.properties.color).to eq(:unknown)
  end
end
