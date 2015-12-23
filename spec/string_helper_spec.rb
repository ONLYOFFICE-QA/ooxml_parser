require 'rspec'
require 'ooxml_parser'

describe OoxmlParser::StringHelper do
  describe 'StringHelper.complex?' do
    it 'StringHelper.complex? check true' do
      expect(OoxmlParser::StringHelper.complex?('3+5i')).to be_truthy
    end

    it 'StringHelper.complex? check false' do
      expect(OoxmlParser::StringHelper.complex?('a+5i')).to be_falsey
    end
  end
end
