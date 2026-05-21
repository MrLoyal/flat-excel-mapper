package com.github.mrloyal;

import com.github.mrloyal.flatexcelmapper.annotation.ExcelColumn;
import com.github.mrloyal.flatexcelmapper.annotation.ExcelEntity;

@ExcelEntity(dataStartRow = 2)
public class Movie {
    private String title;
    private String genre;
    private int releaseYear;
    private String director;
    private String imdbRating;
    private String budget;
    private String boxOffice;

    @ExcelColumn(name = "A")
    public String getTitle() { return title; }
    public void setTitle(String title) { this.title = title; }

    @ExcelColumn(name = "B")
    public String getGenre() { return genre; }
    public void setGenre(String genre) { this.genre = genre; }

    @ExcelColumn(name = "C")
    public int getReleaseYear() { return releaseYear; }
    public void setReleaseYear(int releaseYear) { this.releaseYear = releaseYear; }

    @ExcelColumn(name = "D")
    public String getDirector() { return director; }
    public void setDirector(String director) { this.director = director; }

    @ExcelColumn(name = "E")
    public String getImdbRating() { return imdbRating; }
    public void setImdbRating(String imdbRating) { this.imdbRating = imdbRating; }

    @ExcelColumn(name = "F")
    public String getBudget() { return budget; }
    public void setBudget(String budget) { this.budget = budget; }

    @ExcelColumn(name = "G")
    public String getBoxOffice() { return boxOffice; }
    public void setBoxOffice(String boxOffice) { this.boxOffice = boxOffice; }
}
