require 'spec_helper'

describe 'My behaviour' do
  it 'ImageSize' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/image_size.xlsx')
    expect(xlsx.worksheets.first.drawings).not_to be_empty
    image_path = xlsx.worksheets.first.drawings.first.picture.path_to_image
    expect(File.size?(image_path)).to be > 0
  end

  it 'Incorrect image resource' do
    expect { OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/incorrect_image_resource.xlsx') }.to raise_error(LoadError)
  end

  it 'image_incorrect_link.xlsx' do
    expect { OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/image_incorrect_link.xlsx') }.to raise_error(LoadError)
  end
end
