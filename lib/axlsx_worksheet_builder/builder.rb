# frozen_string_literal: true

module AxlsxWorksheetBuilder
  ##
  # Builder instance
  # @param
  #   worksheet [Axlsx::Worksheet] The worksheet to be built
  class Builder
    attr_accessor :column_headers, :value_getters, :column_types, :worksheet

    def initialize(worksheet)
      @column_headers = []
      @value_getters = []
      @iterated_properties = []
      @column_types = []
      @worksheet = worksheet
    end

    ##
    # Adds a new column to the worksheet
    # @param
    #   column_header [String] The header of the column
    #   &bloc [Proc] A proc that will be called to get the value of the column for each row
    #                This proc will receive as argument each relevant object for the row, see iterate_through_properties
    #   property [Symbol] The property of the object to be added to the column for each row
    #   type [Symbol] The type for the column (:string, :float, :integer, :date, :time or :boolean)
    # @example
    #   add_column("Author name", property: :name)
    #   add_column("Author name") { |author| author.name }
    #   add_column("Author name") do |author|
    #     if author.name.present?
    #       author.name
    #     else
    #       "Anonymous"
    #     end
    #   end
    #   # In case of a nested iteration(iterate_through_properties),
    #   # the values of the properties being iterated through get passed to the proc
    #   add_column("Book title") |author, book| { books.title }
    def add_column(column_header, property: nil, type: nil, &block)
      column_headers << column_header
      column_types << type
      value_getters << ({
        getter_proc: block,
        property: property
      })
    end

    def build_sheet(data)
      worksheet.add_row column_headers
      data.each do |entry|
        add_entry_rows(entry, @iterated_properties, [entry])
      end
    end

    ##
    # Iterates through the property of each object in data
    # @param
    #   property [Symbol] The property of the object to be iterated through
    # @example
    #   # Given a data representing a list of authors with their books:
    #   # [{name: "Martin", books: [{title: "Book 1"}, {title: "Book 2"}]}]
    #   # the following code would create a row for each book of each author
    #   iterate_through_properties(:books)
    def iterate_through_property(property)
      @iterated_properties = [property]
    end

    ##
    # Iterates through the properties of each object in data
    # The proc passed to add_column will receive as argument each objects being iterated through
    # @param
    #   properties [Array<Symbol>] The properties of the object to be iterated through
    # @example
    #   # Given a data representing a list of authors with their books and the chapter of each book:
    #   # [{name: "Martin", books: [{title: "Book 1", chapter: ["Chapter 1", "Chapter 2"]}]}]
    #   # the following code would create a row for each chapter of each book of each author
    #   iterate_through_properties(:books, :chapter)
    #   # In the previous case, the add_column proc would receive the following arguments:
    #   # (author, book, chapter)
    def iterate_through_properties(*properties)
      @iterated_properties = properties
    end

    private

    def add_entry_rows(entry, local_iterated_properties, precedent_entries)
      properties = local_iterated_properties.clone
      property = properties.shift
      if property
        process_entries(entry, property, precedent_entries, properties)
      else
        add_worksheet_row(entry, precedent_entries)
      end
    end

    def process_entries(entry, property, precedent_entries, properties)
      entries = entry_poperty(entry, property)
      entries.each do |entry_property|
        prec_entries = precedent_entries.clone

        prec_entries.push(entry_property)
        add_entry_rows(entry_property, properties, prec_entries)
      end
    end

    def add_worksheet_row(entry, precedent_entries)
      row = value_getters.map do |value_getter|
        property = value_getter[:property]
        if !property.nil?
          entry_poperty(entry, property)
        else
          arguments = precedent_entries[0..(value_getter[:getter_proc].parameters.count - 1)]
          value_getter[:getter_proc].call(*arguments)
        end
      end
      worksheet.add_row row, types: column_types
    end

    def entry_poperty(entry, property)
      entry.send(property)
    rescue StandardError
      entry[property]
    end
  end
end
