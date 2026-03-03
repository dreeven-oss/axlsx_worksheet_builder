# frozen_string_literal: true

require "axlsx"
require_relative "axlsx_worksheet_builder/version"
require_relative "axlsx_worksheet_builder/builder"

##
# Axlsx Worksheet Builder
module AxlsxWorksheetBuilder
  class Error < StandardError; end
  class InvalidWorksheet < ArgumentError; end

  ##
  # Builds the worksheet row, given the columns being added in the block
  # @param
  #   worksheet [Axlsx::Worksheet] The worksheet to be built
  #   data [Array] The data to be added to the worksheet
  # @example
  #   # The following code will build the worksheet with a row for each book of each author
  #   authors = [{name: "Martin", books: [{title: "Book 1"}, {title: "Book 2"}]}]
  #   AxlsxWorksheetBuilder::build(sheet, authors) do |worksheet|
  #     worksheet.iterate_through_property(:books)
  #     worksheet.add_column("Author name", property: :name)
  #     worksheet.add_column("Number of books") { |author| author.books.count }
  #     worksheet.add_column("Book title") { |author, book| books.title }
  #   end
  def self.build(worksheet, data)
    raise InvalidWorksheet, "Worksheet must be an Axlsx::Worksheet" unless worksheet.is_a?(Axlsx::Worksheet)

    builder = Builder.new(worksheet)
    yield builder
    builder.build_sheet(data)
  end
end
