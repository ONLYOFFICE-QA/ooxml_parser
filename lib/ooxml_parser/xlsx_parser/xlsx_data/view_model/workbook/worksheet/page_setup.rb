module OoxmlParser
  # Class for parsing `pageSetup` tag
  class PageSetup < OOXMLDocumentObject
    # @return [Integer] id of paper size
    attr_reader :paper_size
    # @return [Symbol] orientation of page
    attr_reader :orientation

    # Parse PageSetup object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [PageSetup] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'paperSize'
          @paper_size = value.value.to_i
        when 'orientation'
          @orientation = value_to_symbol(value)
        end
      end
      self
    end

    # @return [String] name of paper size
    def paper_size_name
      paper_size_names_list[@paper_size]
    end

    private

    # @return [Array <String>] list of paper size names
    # https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.pagesetup.aspx
    def paper_size_names_list
      ['Default',
       'US letter',
       'US Letter small',
       'Tabloid',
       'Ledger',
       'US Legal',
       'Statement',
       'Executive',
       'A3',
       'A4',
       'A4 small',
       'A5',
       'B4',
       'B5',
       'Folio',
       'Quarto',
       'Standard',
       'Standard',
       'Note',
       'Envelope #9',
       'Envelope #10',
       'Envelope #11',
       'Envelope #12',
       'Envelope #14',
       'C paper',
       'D paper',
       'E paper',
       'Envelope DL',
       'Envelope C5',
       'Envelope C3',
       'Envelope C4',
       'Envelope C6',
       'Envelope C65',
       'Envelope B4',
       'Envelope B5',
       'Envelope B6',
       'Italy envelope',
       'Monarch envelope',
       '6 3/4 envelope',
       'US standard fanfold',
       'German standard fanfold',
       'German legal fanfold',
       'B4',
       'Japanese double postcard ',
       'Standard',
       'Standard',
       'Standard',
       'Invite envelope',
       'Letter oversize',
       'Legal oversize',
       'Tabloid oversize',
       'A4 oversize',
       'Letter transverse paper',
       'A4 transverse paper',
       'Letter extra transverse paper',
       'SuperA/SuperA/A4 paper',
       'SuperB/SuperB/A3 paper',
       'Letter plus paper',
       'A4 plus paper',
       'A5 transverse paper',
       'JIS B5 transverse paper',
       'A3 oversize',
       'A5 oversize ',
       'B5 oversize',
       'A2 paper',
       'A3 transverse paper',
       'A3 extra transverse paper']
    end
  end
end
