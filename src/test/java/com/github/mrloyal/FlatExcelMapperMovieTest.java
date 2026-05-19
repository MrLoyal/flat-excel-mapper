package com.github.mrloyal;

import com.github.mrloyal.flatexcelmapper.FlatExcelMapper;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class FlatExcelMapperMovieTest {

    private static List<Movie> movies;

    @BeforeAll
    static void loadMovies() throws Exception {
        String path = FlatExcelMapperMovieTest.class.getClassLoader().getResource("movies.xlsx").getPath();
        movies = new FlatExcelMapper().read(path, 0, Movie.class);
    }

    @Test
    void readReturnsCorrectCount() {
        assertEquals(12, movies.size());
    }

    @Test
    void readMapsFirstMovieCorrectly() {
        Movie m = movies.get(0);
        assertEquals("The Shawshank Redemption", m.getTitle());
        assertEquals("Drama", m.getGenre());
        assertEquals(1994, m.getReleaseYear());
        assertEquals("Frank Darabont", m.getDirector());
        assertEquals("9.3", m.getImdbRating());
    }

    @Test
    void readMapsReleaseYearAsInt() {
        assertEquals(1972, movies.get(8).getReleaseYear()); // The Godfather
        assertEquals(2023, movies.get(10).getReleaseYear()); // Oppenheimer
    }

    // Column E uses a custom "d.m" date format to encode ratings as date serials.
    // DataFormatter resolves them back to "day.month" decimal strings (e.g. "9.3").
    @Test
    void readMapsDateFormattedRating() {
        assertEquals("9.3", movies.get(0).getImdbRating());  // Shawshank — date-formatted cell
        assertEquals("9.0", movies.get(1).getImdbRating());  // Dark Knight — stored as plain string
        assertEquals("9.2", movies.get(8).getImdbRating());  // Godfather
    }

    @Test
    void readCountsNolanMovies() {
        long count = movies.stream()
                .filter(m -> "Christopher Nolan".equals(m.getDirector()))
                .count();
        assertEquals(4, count); // Dark Knight, Inception, Interstellar, Oppenheimer
    }

    @Test
    void readFindsLowestRatedMovie() {
        Movie lowestRated = movies.stream()
                .min((a, b) -> Double.compare(
                        Double.parseDouble(a.getImdbRating()),
                        Double.parseDouble(b.getImdbRating())))
                .orElseThrow();
        assertEquals("The Lion King", lowestRated.getTitle());
        assertEquals("6.8", lowestRated.getImdbRating());
    }

    @Test
    void readFindsHighestGrossingMovie() {
        Movie topGrosser = movies.stream()
                .max((a, b) -> Double.compare(
                        Double.parseDouble(a.getBoxOffice()),
                        Double.parseDouble(b.getBoxOffice())))
                .orElseThrow();
        assertEquals("Avengers: Endgame", topGrosser.getTitle());
        assertEquals("2798", topGrosser.getBoxOffice());
    }

    // Pulp Fiction's budget cell is also date-formatted ("d.m"), encoding $8.5M as May 8.
    @Test
    void readVerifiesPulpFictionBudget() {
        Movie pulpFiction = movies.get(3);
        assertEquals("Pulp Fiction", pulpFiction.getTitle());
        assertEquals("8.5", pulpFiction.getBudget());
    }

    @Test
    void readVerifiesDirectorSpotCheck() {
        assertEquals("Quentin Tarantino", movies.get(3).getDirector()); // Pulp Fiction
        assertEquals("David Fincher", movies.get(7).getDirector());     // Fight Club
        assertEquals("Robert Zemeckis", movies.get(4).getDirector());   // Forrest Gump
    }
}
